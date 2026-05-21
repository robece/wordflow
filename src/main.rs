use std::path::PathBuf;

use anyhow::Result;
use clap::{Parser, Subcommand};
use serde_json;
use wordflow::{
    AnchorMatchMode, AnchorSpec, AnchorTarget, AutomationSpec, CommentSpec, CommentUpdate,
    add_comment_to_docx, apply_spec_file_to_docx, apply_spec_file_to_docx_dry_run,
    delete_docx_comment, diff_docx_files, find_anchors_in_docx,
    inspect_normalization_file, list_docx_parts, migrate_source_to_ooxml, normalize_docx_file,
    list_docx_comments, prepare_work_session, publish_session_to_next_version,
    publish_spec_file_to_docx, publish_spec_file_to_next_version, update_docx_comment,
    validate_docx_file, validate_source_fidelity_file, validate_spec_file, ParagraphStyle,
    PublishTargetMode, reconstruct_session, show_session, track_session,
};

#[derive(Parser, Debug)]
#[command(
    author,
    version,
    about = "Reliable .docx automation through OpenXML, without Word COM."
)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand, Debug)]
enum Commands {
    /// Insert paragraphs after anchor text based on a JSON spec file.
    Insert {
        /// Input .docx file.
        #[arg(long)]
        input: PathBuf,

        /// Output .docx file.
        #[arg(long)]
        output: PathBuf,

        /// JSON spec describing anchors and paragraph entries.
        #[arg(long)]
        spec: PathBuf,

        /// Validate and print the result without writing the output file.
        #[arg(long, default_value_t = false)]
        dry_run: bool,
    },
    /// Stage the source in a temp workspace, validate a candidate, then publish only on success.
    Publish {
        /// Input .docx file.
        #[arg(long)]
        input: PathBuf,

        /// Final output .docx file.
        #[arg(long)]
        output: PathBuf,

        /// JSON spec describing anchors and paragraph entries.
        #[arg(long)]
        spec: PathBuf,

        /// Temp root used for staging and candidate generation.
        #[arg(long, default_value = r"C:\Temp\wordflow")]
        temp_root: PathBuf,
    },
    /// Publish to the next detected versioned filename (for example, v014 -> v015).
    PublishNext {
        /// Input .docx file when publishing without a persisted work session.
        #[arg(long)]
        input: Option<PathBuf>,

        /// Reusable work-session identifier created by prepare-session.
        #[arg(long)]
        session: Option<String>,

        /// JSON spec describing anchors and paragraph entries.
        #[arg(long)]
        spec: PathBuf,

        /// Optional override for the final output directory.
        #[arg(long)]
        output_dir: Option<PathBuf>,

        /// Whether to create a new version or continue working on the latest published one.
        #[arg(long, default_value = "next-version")]
        target: String,

        /// Temp root used for staging, session state, and candidate generation.
        #[arg(long, default_value = r"C:\Temp\wordflow")]
        temp_root: PathBuf,
    },
    /// Prepare or reuse a normalized working copy inside a reusable temp session.
    PrepareSession {
        /// Input source document.
        #[arg(long)]
        input: PathBuf,

        /// Session identifier. Reusing the same id reuses the cache when the source is unchanged.
        #[arg(long)]
        session: String,

        /// Temp root used for session state and normalized working copies.
        #[arg(long, default_value = r"C:\Temp\wordflow")]
        temp_root: PathBuf,

        /// Optional trusted OOXML reference used for fidelity comparison.
        #[arg(long)]
        trusted_ooxml: Option<PathBuf>,

        /// Allow the guarded Word COM fallback, but only for detected EncryptedPackage sources.
        #[arg(long, default_value_t = false)]
        allow_word_com_encrypted_package: bool,
    },
    /// Validate a JSON spec before any document edits run.
    ValidateSpec {
        #[arg(long)]
        spec: PathBuf,
    },
    /// Add a review comment after an anchor.
    AddComment {
        #[arg(long)]
        input: PathBuf,
        #[arg(long)]
        output: PathBuf,
        #[arg(long)]
        anchor: String,
        #[arg(long, default_value = "contains")]
        mode: String,
        #[arg(long, default_value_t = 1)]
        occurrence: usize,
        #[arg(long)]
        part: Option<String>,
        #[arg(long, default_value = "GitHub Copilot CLI review comment")]
        text: String,
        #[arg(long)]
        comment_text: String,
        #[arg(long, default_value = "GitHub Copilot CLI")]
        author: String,
        #[arg(long, default_value = "GCC")]
        initials: String,
        #[arg(long)]
        date: Option<String>,
        #[arg(long, default_value = "yellow")]
        highlight: String,
    },
    /// List comments inside a normalized OOXML .docx.
    ListComments {
        #[arg(long)]
        input: PathBuf,
    },
    /// Update the metadata or body of an existing comment.
    UpdateComment {
        #[arg(long)]
        input: PathBuf,
        #[arg(long)]
        output: PathBuf,
        #[arg(long)]
        id: u32,
        #[arg(long)]
        comment_text: Option<String>,
        #[arg(long)]
        author: Option<String>,
        #[arg(long)]
        initials: Option<String>,
        #[arg(long)]
        clear_initials: bool,
        #[arg(long)]
        date: Option<String>,
        #[arg(long)]
        clear_date: bool,
        #[arg(long)]
        highlight: Option<String>,
    },
    /// Delete an existing comment and its in-document references.
    DeleteComment {
        #[arg(long)]
        input: PathBuf,
        #[arg(long)]
        output: PathBuf,
        #[arg(long)]
        id: u32,
    },
    /// Convert a protected or non-OOXML source into a validated OOXML .docx through Word.
    Migrate {
        /// Input source document.
        #[arg(long)]
        input: PathBuf,

        /// Final migrated OOXML .docx output.
        #[arg(long)]
        output: PathBuf,

        /// Temp root used for staging and migration artifacts.
        #[arg(long, default_value = r"C:\Temp\wordflow")]
        temp_root: PathBuf,

        /// Optional trusted OOXML reference used for fidelity comparison.
        #[arg(long)]
        trusted_ooxml: Option<PathBuf>,
    },
    /// Detect whether a source is already normalized and publish an OOXML-safe copy when possible.
    Normalize {
        /// Input source document.
        #[arg(long)]
        input: PathBuf,

        /// Output normalized OOXML .docx.
        #[arg(long)]
        output: PathBuf,

        /// Temp root reserved for future normalization backends.
        #[arg(long, default_value = r"C:\Temp\wordflow")]
        temp_root: PathBuf,

        /// Optional trusted OOXML reference used for fidelity comparison.
        #[arg(long)]
        trusted_ooxml: Option<PathBuf>,

        /// Allow the guarded Word COM fallback, but only for detected EncryptedPackage sources.
        #[arg(long, default_value_t = false)]
        allow_word_com_encrypted_package: bool,
    },
    /// Inspect high-level document information and validation status.
    Inspect {
        #[arg(long)]
        input: PathBuf,
    },
    /// List all parts inside the .docx package.
    ListParts {
        #[arg(long)]
        input: PathBuf,
    },
    /// Search paragraph anchors in one part or across the document package.
    FindAnchors {
        #[arg(long)]
        input: PathBuf,
        #[arg(long)]
        text: String,
        #[arg(long, default_value = "contains")]
        mode: String,
        #[arg(long, default_value_t = 1)]
        occurrence: usize,
        #[arg(long)]
        part: Option<String>,
    },
    /// Validate the OpenXML package structure and XML parsing.
    Validate {
        #[arg(long)]
        input: PathBuf,
    },
    /// Compare two OOXML documents and report whether inherited formatting was preserved.
    CheckFidelity {
        #[arg(long)]
        before: PathBuf,
        #[arg(long)]
        after: PathBuf,
        /// Optional spec to ignore paragraphs intentionally changed by the requested operations.
        #[arg(long)]
        spec: Option<PathBuf>,
    },
    /// Show changed parts between two .docx files.
    DiffDocx {
        #[arg(long)]
        before: PathBuf,
        #[arg(long)]
        after: PathBuf,
    },
    /// Print the session history for a document as JSON.
    /// Accepts either a .docx path (derives the session file) or a .session.json path directly.
    ShowSession {
        #[arg(long)]
        input: PathBuf,
    },
    /// Infer a session history from sequential versioned .docx files in a folder.
    /// Merges into any existing session file for the same document.
    ReconstructSession {
        /// Folder containing the versioned .docx files.
        #[arg(long)]
        folder: PathBuf,
        /// Document stem without version suffix (e.g. "report").
        #[arg(long)]
        document: String,
    },
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    match cli.command {
        Commands::Insert {
            input,
            output,
            spec,
            dry_run,
        } => {
            let parsed = AutomationSpec::from_path(&spec)?;
            let spec_base = spec.parent().unwrap_or(std::path::Path::new("."));
            if dry_run {
                let report = apply_spec_file_to_docx_dry_run(&input, &parsed, spec_base)?;
                println!(
                    "Dry run complete: {} XML parts checked, {} issues",
                    report.xml_parts_checked,
                    report.issues.len()
                );
                for issue in report.issues {
                    println!("ISSUE\t{}", issue);
                }
            } else {
                apply_spec_file_to_docx(&input, &output, &parsed, spec_base)?;
                println!("Updated {}", output.display());
            }
        }
        Commands::Publish {
            input,
            output,
            spec,
            temp_root,
        } => {
            let parsed = AutomationSpec::from_path(&spec)?;
            let spec_base = spec.parent().unwrap_or(std::path::Path::new("."));
            let report =
                publish_spec_file_to_docx(&input, &output, &parsed, spec_base, Some(&temp_root))?;
            println!(
                "Published {}\tXML checked\t{}",
                report.published_output.display(),
                report.xml_parts_checked
            );
            if let Ok(session_path) = track_session(
                &report.published_output,
                "publish",
                Some(&input),
                Some(&spec),
                "ok",
                None,
            ) {
                println!("Session\t{}", session_path.display());
            }
        }
        Commands::PublishNext {
            input,
            session,
            spec,
            output_dir,
            target,
            temp_root,
        } => {
            let parsed = AutomationSpec::from_path(&spec)?;
            let spec_base = spec.parent().unwrap_or(std::path::Path::new("."));
            let target_mode = parse_publish_target_mode(&target)?;
            let report = match (input, session) {
                (Some(input), None) => publish_spec_file_to_next_version(
                    &input,
                    &parsed,
                    spec_base,
                    Some(&temp_root),
                    output_dir.as_deref(),
                    target_mode,
                )?,
                (None, Some(session_id)) => publish_session_to_next_version(
                    &session_id,
                    &parsed,
                    spec_base,
                    Some(&temp_root),
                    output_dir.as_deref(),
                    target_mode,
                )?,
                (Some(_), Some(_)) => {
                    anyhow::bail!("publish-next accepts either --input or --session, but not both")
                }
                (None, None) => anyhow::bail!("publish-next requires either --input or --session"),
            };
            println!(
                "Published {}\tMode\t{}\tVersion\tv{:03}\tXML checked\t{}",
                report.published_output.display(),
                report.mode.as_str(),
                report.version_number,
                report.xml_parts_checked
            );
            if let Ok(session_path) = track_session(
                &report.published_output,
                "publish-next",
                None,
                Some(&spec),
                "ok",
                Some(&format!("mode={} version=v{:03}", report.mode.as_str(), report.version_number)),
            ) {
                println!("Session\t{}", session_path.display());
            }
        }
        Commands::PrepareSession {
            input,
            session,
            temp_root,
            trusted_ooxml,
            allow_word_com_encrypted_package,
        } => {
            let report = prepare_work_session(
                &input,
                &session,
                Some(&temp_root),
                trusted_ooxml.as_deref(),
                allow_word_com_encrypted_package,
            )?;
            println!(
                "Session\t{}\tCache hit\t{}\tFormat\t{}\tWorking copy\t{}",
                report.session_id,
                report.cache_hit,
                report.detected_format.as_str(),
                report.normalized_input.display()
            );
        }
        Commands::ValidateSpec { spec } => {
            let operations = validate_spec_file(&spec)?;
            println!("Valid\ttrue");
            println!("Operations\t{operations}");
        }
        Commands::AddComment {
            input,
            output,
            anchor,
            mode,
            occurrence,
            part,
            text,
            comment_text,
            author,
            initials,
            date,
            highlight,
        } => {
            let anchor = AnchorTarget::Structured(AnchorSpec {
                text: anchor,
                mode: parse_anchor_mode(&mode)?,
                occurrence,
            });
            let comment = CommentSpec {
                text,
                comment_text,
                author,
                initials: Some(initials),
                date,
                style: ParagraphStyle::Normal,
                highlight,
            };
            add_comment_to_docx(&input, &output, part.as_deref(), &anchor, &comment)?;
            println!("Updated {}", output.display());
        }
        Commands::ListComments { input } => {
            for comment in list_docx_comments(&input)? {
                let locations = comment
                    .locations
                    .iter()
                    .map(|location| format!("{}:{}", location.part, location.paragraph_text))
                    .collect::<Vec<_>>()
                    .join(" | ");
                println!(
                    "{}\t{}\t{}\t{}\t{}\t{}\t{}",
                    comment.id,
                    comment.author,
                    comment.initials.unwrap_or_default(),
                    comment.date.unwrap_or_default(),
                    comment.highlight,
                    comment.comment_text,
                    locations
                );
            }
        }
        Commands::UpdateComment {
            input,
            output,
            id,
            comment_text,
            author,
            initials,
            clear_initials,
            date,
            clear_date,
            highlight,
        } => {
            let update = CommentUpdate {
                comment_text,
                author,
                initials: if clear_initials {
                    Some(None)
                } else {
                    initials.map(Some)
                },
                date: if clear_date { Some(None) } else { date.map(Some) },
                highlight,
            };
            update_docx_comment(&input, &output, id, &update)?;
            println!("Updated {}", output.display());
        }
        Commands::DeleteComment { input, output, id } => {
            delete_docx_comment(&input, &output, id)?;
            println!("Updated {}", output.display());
        }
        Commands::Migrate {
            input,
            output,
            temp_root,
            trusted_ooxml,
        } => {
            let report = migrate_source_to_ooxml(
                &input,
                &output,
                Some(&temp_root),
                trusted_ooxml.as_deref(),
            )?;
            println!(
                "Migrated {}\tXML checked\t{}\tText exports match\t{}\tWord fidelity match\t{}",
                report.published_output.display(),
                report.xml_parts_checked,
                report.text_exports_match
                ,
                report.word_fidelity_match
            );
        }
        Commands::Normalize {
            input,
            output,
            temp_root,
            trusted_ooxml,
            allow_word_com_encrypted_package,
        } => {
            let report = normalize_docx_file(
                &input,
                &output,
                Some(&temp_root),
                trusted_ooxml.as_deref(),
                allow_word_com_encrypted_package,
            )?;
            println!(
                "Normalized {}\tFormat\t{}\tAlready normalized\t{}\tXML checked\t{}",
                report.published_output.display(),
                report.detected_format.as_str(),
                report.already_normalized,
                report.xml_parts_checked
            );
        }
        Commands::Inspect { input } => {
            let normalization = inspect_normalization_file(&input)?;
            println!("Format\t{}", normalization.format.as_str());
            println!("Normalized\t{}", normalization.is_normalized);
            for detail in normalization.details {
                println!("DETAIL\t{}", detail);
            }
            if normalization.is_normalized {
                let parts = list_docx_parts(&input)?;
                let validation = validate_docx_file(&input)?;
                println!("Parts\t{}", parts.len());
                println!("XML checked\t{}", validation.xml_parts_checked);
                println!("Valid\t{}", validation.is_valid());
                for issue in validation.issues {
                    println!("ISSUE\t{}", issue);
                }
            }
        }
        Commands::ListParts { input } => {
            for part in list_docx_parts(&input)? {
                if !part.is_dir {
                    println!(
                        "{}\t{}\t{}",
                        part.name,
                        part.size,
                        part.sha256.unwrap_or_default()
                    );
                }
            }
        }
        Commands::FindAnchors {
            input,
            text,
            mode,
            occurrence,
            part,
        } => {
            let mode = parse_anchor_mode(&mode)?;
            let anchor = AnchorTarget::Structured(AnchorSpec {
                text,
                mode,
                occurrence,
            });
            for found in find_anchors_in_docx(&input, part.as_deref(), &anchor)? {
                println!("{}\t{}\t{}", found.part, found.index, found.text);
            }
        }
        Commands::Validate { input } => {
            let report = validate_docx_file(&input)?;
            println!("XML checked\t{}", report.xml_parts_checked);
            println!("Valid\t{}", report.is_valid());
            for issue in report.issues {
                println!("ISSUE\t{}", issue);
            }
        }
        Commands::CheckFidelity {
            before,
            after,
            spec,
        } => {
            let parsed_spec = match spec {
                Some(path) => Some(AutomationSpec::from_path(&path)?),
                None => None,
            };
            let report =
                validate_source_fidelity_file(&before, &after, parsed_spec.as_ref())?;
            println!("Valid\t{}", report.is_valid());
            for issue in report.issues {
                println!("ISSUE\t{}", issue);
            }
        }
        Commands::DiffDocx { before, after } => {
            let diff = diff_docx_files(&before, &after)?;
            for part in diff.added_parts {
                println!("ADDED\t{}", part);
            }
            for part in diff.removed_parts {
                println!("REMOVED\t{}", part);
            }
            for part in diff.changed_parts {
                println!("CHANGED\t{}", part);
            }
        }
        Commands::ShowSession { input } => {
            let session = show_session(&input)?;
            println!("{}", serde_json::to_string_pretty(&session)?);
        }
        Commands::ReconstructSession { folder, document } => {
            let session = reconstruct_session(&folder, &document)?;
            println!("{}", serde_json::to_string_pretty(&session)?);
            eprintln!(
                "Reconstructed {} entries — saved to {}.session.json",
                session.entries.len(),
                document
            );
        }
    }

    Ok(())
}

fn parse_anchor_mode(value: &str) -> Result<AnchorMatchMode> {
    Ok(match value {
        "contains" => AnchorMatchMode::Contains,
        "equals" => AnchorMatchMode::Equals,
        "starts-with" => AnchorMatchMode::StartsWith,
        "ends-with" => AnchorMatchMode::EndsWith,
        _ => anyhow::bail!("unsupported anchor mode: {value}"),
    })
}

fn parse_publish_target_mode(value: &str) -> Result<PublishTargetMode> {
    Ok(match value {
        "next-version" => PublishTargetMode::NextVersion,
        "latest" => PublishTargetMode::Latest,
        _ => anyhow::bail!("unsupported publish target mode: {value}"),
    })
}
