use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use anyhow::Result;
use regex::Regex;
use serde::{Deserialize, Serialize};

// ── Types ─────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SessionEntry {
    pub timestamp: String,
    pub command: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub input: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub output: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub spec: Option<String>,
    pub result: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SessionFile {
    pub document: String,
    pub session_file: String,
    pub created_at: String,
    pub updated_at: String,
    pub entries: Vec<SessionEntry>,
}

// ── Time helpers (no chrono dependency) ───────────────────────────────────────

fn system_time_to_rfc3339(t: SystemTime) -> String {
    let secs = t.duration_since(UNIX_EPOCH).unwrap_or_default().as_secs();
    epoch_secs_to_rfc3339(secs)
}

fn epoch_secs_to_rfc3339(mut secs: u64) -> String {
    let s = secs % 60; secs /= 60;
    let mi = secs % 60; secs /= 60;
    let h = secs % 24; secs /= 24;
    let mut year = 1970u64;
    loop {
        let dy = if is_leap(year) { 366 } else { 365 };
        if secs < dy { break; }
        secs -= dy;
        year += 1;
    }
    let mdays: [u64; 12] = [31, if is_leap(year) { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let mut month = 0u64;
    for &md in &mdays {
        if secs < md { break; }
        secs -= md;
        month += 1;
    }
    format!("{year:04}-{mo:02}-{d:02}T{h:02}:{mi:02}:{s:02}Z",
        mo = month + 1,
        d  = secs + 1)
}

fn is_leap(y: u64) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

pub fn now_rfc3339() -> String {
    system_time_to_rfc3339(SystemTime::now())
}

pub fn mtime_rfc3339(path: &Path) -> String {
    path.metadata()
        .ok()
        .and_then(|m| m.modified().ok())
        .map(system_time_to_rfc3339)
        .unwrap_or_else(now_rfc3339)
}

// ── Path helpers ──────────────────────────────────────────────────────────────

/// Strip a trailing version suffix (-v015, -v15, -V3, etc.) from a file stem.
pub fn document_stem(path: &Path) -> Option<String> {
    let stem = path.file_stem()?.to_str()?;
    let re = Regex::new(r"-[vV]\d+$").ok()?;
    Some(re.replace(stem, "").to_string())
}

/// Derive `<dir>/.wordflow/<stem>.session.json` from a versioned .docx path.
pub fn resolve_session_path(document_path: &Path) -> Option<PathBuf> {
    let dir = document_path.parent()?;
    let stem = document_stem(document_path)?;
    let wordflow_dir = if dir.as_os_str().is_empty() {
        std::env::current_dir().ok()?.join(".wordflow")
    } else {
        dir.join(".wordflow")
    };
    Some(wordflow_dir.join(format!("{stem}.session.json")))
}

// ── SessionFile impl ──────────────────────────────────────────────────────────

impl SessionFile {
    pub fn load(path: &Path) -> Result<Self> {
        let raw = fs::read_to_string(path)?;
        Ok(serde_json::from_str(&raw)?)
    }

    pub fn load_or_create(path: &Path, doc_stem: &str) -> Result<Self> {
        if path.exists() {
            Self::load(path)
        } else {
            let now = now_rfc3339();
            Ok(Self {
                document: doc_stem.to_string(),
                session_file: path.display().to_string(),
                created_at: now.clone(),
                updated_at: now,
                entries: vec![],
            })
        }
    }

    pub fn append(&mut self, entry: SessionEntry) {
        self.updated_at = now_rfc3339();
        self.entries.push(entry);
    }

    pub fn save(&self, path: &Path) -> Result<()> {
        let json = serde_json::to_string_pretty(self)?;
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)?;
            }
        }
        fs::write(path, json)?;
        Ok(())
    }
}

// ── Public API ────────────────────────────────────────────────────────────────

/// Append one entry to the session file derived from `output_path`.
/// Creates the session file if it does not yet exist.
pub fn track_session(
    output_path: &Path,
    command: &str,
    input: Option<&Path>,
    spec: Option<&Path>,
    result: &str,
    note: Option<&str>,
) -> Result<PathBuf> {
    let session_path = resolve_session_path(output_path).ok_or_else(|| {
        anyhow::anyhow!("cannot resolve session path from {}", output_path.display())
    })?;

    let doc_stem = document_stem(output_path).unwrap_or_else(|| "document".to_string());
    let mut session = SessionFile::load_or_create(&session_path, &doc_stem)?;

    session.append(SessionEntry {
        timestamp: now_rfc3339(),
        command: command.to_string(),
        input: input.map(|p| p.display().to_string()),
        output: Some(output_path.display().to_string()),
        spec: spec.map(|p| p.display().to_string()),
        result: result.to_string(),
        note: note.map(|s| s.to_string()),
    });

    session.save(&session_path)?;
    Ok(session_path)
}

/// Load and return a session file.
/// Accepts either a .docx path (derives the session path) or a .session.json path directly.
pub fn show_session(path: &Path) -> Result<SessionFile> {
    let session_path = if path.extension().and_then(|e| e.to_str()) == Some("json") {
        path.to_path_buf()
    } else {
        resolve_session_path(path)
            .ok_or_else(|| anyhow::anyhow!("cannot resolve session path from {}", path.display()))?
    };
    SessionFile::load(&session_path)
}

/// Infer a session history from sequential versioned .docx files in a folder.
/// Merges into any existing session file for the same document stem.
pub fn reconstruct_session(folder: &Path, document_stem_arg: &str) -> Result<SessionFile> {
    let re_ver = Regex::new(r"[vV](\d+)$")?;

    let mut files: Vec<PathBuf> = fs::read_dir(folder)?
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| {
            let stem = p.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            let ext = p.extension().and_then(|e| e.to_str()).unwrap_or("");
            ext == "docx" && stem.starts_with(document_stem_arg)
        })
        .collect();

    files.sort_by_key(|p| {
        let stem = p.file_stem().and_then(|s| s.to_str()).unwrap_or("");
        re_ver.captures(stem)
            .and_then(|c| c.get(1))
            .and_then(|m| m.as_str().parse::<u32>().ok())
            .unwrap_or(0)
    });

    let session_path = folder.join(".wordflow").join(format!("{document_stem_arg}.session.json"));

    let created_at = files.first()
        .map(|p| mtime_rfc3339(p))
        .unwrap_or_else(now_rfc3339);

    let mut session = if session_path.exists() {
        SessionFile::load(&session_path)?
    } else {
        SessionFile {
            document: document_stem_arg.to_string(),
            session_file: session_path.display().to_string(),
            created_at,
            updated_at: now_rfc3339(),
            entries: vec![],
        }
    };

    for pair in files.windows(2) {
        let before = &pair[0];
        let after  = &pair[1];

        let already_tracked = session.entries.iter().any(|e| {
            e.output.as_deref() == Some(&after.display().to_string())
                && e.command == "reconstructed"
        });
        if already_tracked {
            continue;
        }

        let diff = crate::diff_docx_files(before, after)?;
        let changed = diff.changed_parts.len() + diff.added_parts.len() + diff.removed_parts.len();

        session.entries.push(SessionEntry {
            timestamp: mtime_rfc3339(after),
            command: "reconstructed".to_string(),
            input: Some(before.display().to_string()),
            output: Some(after.display().to_string()),
            spec: None,
            result: "ok".to_string(),
            note: Some(format!(
                "{changed} changed XML part(s) — inferred from diff, spec and intent not available"
            )),
        });
    }

    session.entries.sort_by(|a, b| a.timestamp.cmp(&b.timestamp));
    session.updated_at = now_rfc3339();
    session.save(&session_path)?;

    Ok(session)
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn document_stem_strips_version_suffix() {
        assert_eq!(document_stem(Path::new("report-v007.docx")).unwrap(), "report");
        assert_eq!(document_stem(Path::new("doc-v1.docx")).unwrap(), "doc");
        assert_eq!(document_stem(Path::new("strategy-V003.docx")).unwrap(), "strategy");
    }

    #[test]
    fn document_stem_returns_full_stem_without_version() {
        assert_eq!(document_stem(Path::new("report.docx")).unwrap(), "report");
        assert_eq!(document_stem(Path::new("my-document.docx")).unwrap(), "my-document");
    }

    #[test]
    fn resolve_session_path_derives_session_file() {
        let path = Path::new("/folder/report-v007.docx");
        let session = resolve_session_path(path).unwrap();
        let s = session.to_string_lossy();
        assert!(s.contains(".wordflow"), "expected .wordflow in path: {s}");
        assert!(s.ends_with("report.session.json"), "expected report.session.json: {s}");
    }

    #[test]
    fn resolve_session_path_works_without_version() {
        let path = Path::new("/folder/report.docx");
        let session = resolve_session_path(path).unwrap();
        let s = session.to_string_lossy();
        assert!(s.contains(".wordflow"), "expected .wordflow in path: {s}");
        assert!(s.ends_with("report.session.json"), "expected report.session.json: {s}");
    }

    #[test]
    fn epoch_secs_to_rfc3339_formats_unix_epoch() {
        assert_eq!(epoch_secs_to_rfc3339(0), "1970-01-01T00:00:00Z");
    }

    #[test]
    fn epoch_secs_to_rfc3339_formats_known_date() {
        assert_eq!(epoch_secs_to_rfc3339(1735689600), "2025-01-01T00:00:00Z");
    }

    #[test]
    fn is_leap_identifies_leap_years() {
        assert!(is_leap(2000));
        assert!(is_leap(2024));
        assert!(!is_leap(1900));
        assert!(!is_leap(2023));
    }

    #[test]
    fn session_file_create_load_save_roundtrip() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("test.session.json");

        let mut session = SessionFile::load_or_create(&path, "test-doc").unwrap();
        assert_eq!(session.document, "test-doc");
        assert!(session.entries.is_empty());

        session.append(SessionEntry {
            timestamp: "2026-01-01T00:00:00Z".to_string(),
            command: "publish".to_string(),
            input: Some("input.docx".to_string()),
            output: Some("output.docx".to_string()),
            spec: Some("spec.json".to_string()),
            result: "ok".to_string(),
            note: None,
        });
        session.save(&path).unwrap();

        let loaded = SessionFile::load(&path).unwrap();
        assert_eq!(loaded.entries.len(), 1);
        assert_eq!(loaded.entries[0].command, "publish");
        assert_eq!(loaded.document, "test-doc");
    }

    #[test]
    fn track_session_creates_and_appends_entries() {
        let dir = tempdir().unwrap();
        let output = dir.path().join("report-v007.docx");
        let input = dir.path().join("report-v006.docx");
        let spec = dir.path().join("spec.json");

        let session_path =
            track_session(&output, "publish", Some(&input), Some(&spec), "ok", None).unwrap();
        assert!(session_path.exists());

        let session = SessionFile::load(&session_path).unwrap();
        assert_eq!(session.entries.len(), 1);
        assert_eq!(session.entries[0].command, "publish");
        assert_eq!(session.entries[0].result, "ok");
    }

    #[test]
    fn track_session_appends_multiple_entries() {
        let dir = tempdir().unwrap();
        let output = dir.path().join("report-v007.docx");

        track_session(&output, "first", None, None, "ok", None).unwrap();
        track_session(&output, "second", None, None, "ok", Some("a note")).unwrap();

        let session_path = resolve_session_path(&output).unwrap();
        let session = SessionFile::load(&session_path).unwrap();
        assert_eq!(session.entries.len(), 2);
        assert_eq!(session.entries[0].command, "first");
        assert_eq!(session.entries[1].command, "second");
        assert_eq!(session.entries[1].note.as_deref(), Some("a note"));
    }
}
