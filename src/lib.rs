pub mod session;
pub use session::{
    SessionEntry, SessionFile, reconstruct_session, show_session, track_session,
};

use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::process::{self, Command};
use std::time::{SystemTime, UNIX_EPOCH};

use anyhow::{Context, Result, bail};
use regex::Regex;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use xmltree::{Element, EmitterConfig, Namespace, XMLNode};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const DOCUMENT_XML_PATH: &str = "word/document.xml";
const CONTENT_TYPES_XML_PATH: &str = "[Content_Types].xml";
const ROOT_RELS_XML_PATH: &str = "_rels/.rels";
const DEFAULT_HIGHLIGHT: &str = "green";
const ZIP_HEADER: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];
const OLE_HEADER: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W15_NS: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CID_NS: &str = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const REL_TYPE_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_TYPE_CORE_PROPS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_TYPE_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_TYPE_COMMENTS_EXTENDED: &str =
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const REL_TYPE_COMMENTS_IDS: &str =
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const REL_TYPE_PEOPLE: &str = "http://schemas.microsoft.com/office/2011/relationships/people";
const REL_TYPE_FOOTNOTES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const REL_TYPE_ENDNOTES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
const REL_TYPE_HEADER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const REL_TYPE_FOOTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const CORE_PROPS_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.core-properties+xml";
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXTENDED_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
const COMMENTS_IDS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
const PEOPLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml";
const FOOTNOTES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const ENDNOTES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const DEFAULT_TEMP_ROOT: &str = r"C:\Temp\wordflow";
const SESSION_METADATA_FILE: &str = "session.json";
const SESSION_WORKING_DOCX_FILE: &str = "working-normalized.docx";
const SUPPORTED_HIGHLIGHTS: &[&str] = &[
    "black",
    "blue",
    "cyan",
    "darkBlue",
    "darkCyan",
    "darkGray",
    "darkGreen",
    "darkMagenta",
    "darkRed",
    "darkYellow",
    "green",
    "lightGray",
    "magenta",
    "none",
    "red",
    "white",
    "yellow",
];

#[derive(Debug, Clone, Deserialize)]
pub struct AutomationSpec {
    pub operations: Vec<Operation>,
}

#[derive(Debug, Clone)]
pub struct ValidationReport {
    pub xml_parts_checked: usize,
    pub issues: Vec<String>,
}

impl ValidationReport {
    pub fn is_valid(&self) -> bool {
        self.issues.is_empty()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum DocumentFormat {
    OoxmlZip,
    OleEncryptedPackage,
    OleWordBinary,
    OleCompound,
    Unknown,
}

impl DocumentFormat {
    pub fn as_str(&self) -> &'static str {
        match self {
            DocumentFormat::OoxmlZip => "ooxml-zip",
            DocumentFormat::OleEncryptedPackage => "ole-encrypted-package",
            DocumentFormat::OleWordBinary => "ole-word-binary",
            DocumentFormat::OleCompound => "ole-compound",
            DocumentFormat::Unknown => "unknown",
        }
    }
}

#[derive(Debug, Clone)]
pub struct NormalizationReport {
    pub format: DocumentFormat,
    pub is_normalized: bool,
    pub requires_normalization: bool,
    pub details: Vec<String>,
}

#[derive(Debug, Clone)]
pub struct NormalizeWorkflowReport {
    pub detected_format: DocumentFormat,
    pub already_normalized: bool,
    pub xml_parts_checked: usize,
    pub published_output: PathBuf,
}

#[derive(Debug, Clone)]
pub struct AnchorMatch {
    pub part: String,
    pub index: usize,
    pub text: String,
}

#[derive(Debug, Clone)]
pub struct PartSummary {
    pub name: String,
    pub is_dir: bool,
    pub size: usize,
    pub sha256: Option<String>,
}

#[derive(Debug, Clone)]
pub struct DiffSummary {
    pub added_parts: Vec<String>,
    pub removed_parts: Vec<String>,
    pub changed_parts: Vec<String>,
}

#[derive(Debug, Clone)]
pub struct PublishWorkflowReport {
    pub xml_parts_checked: usize,
    pub published_output: PathBuf,
}

#[derive(Debug, Clone)]
pub struct PublishNextWorkflowReport {
    pub xml_parts_checked: usize,
    pub published_output: PathBuf,
    pub version_number: usize,
    pub mode: PublishTargetMode,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PublishTargetMode {
    NextVersion,
    Latest,
}

impl PublishTargetMode {
    pub fn as_str(&self) -> &'static str {
        match self {
            PublishTargetMode::NextVersion => "next-version",
            PublishTargetMode::Latest => "latest",
        }
    }
}

#[derive(Debug, Clone)]
pub struct WorkSessionReport {
    pub session_id: String,
    pub session_dir: PathBuf,
    pub normalized_input: PathBuf,
    pub detected_format: DocumentFormat,
    pub cache_hit: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentLocation {
    pub part: String,
    pub paragraph_text: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentRecord {
    pub id: u32,
    pub author: String,
    pub initials: Option<String>,
    pub date: Option<String>,
    pub highlight: String,
    pub comment_text: String,
    pub locations: Vec<CommentLocation>,
}

#[derive(Debug, Clone, Default)]
pub struct CommentUpdate {
    pub comment_text: Option<String>,
    pub author: Option<String>,
    pub initials: Option<Option<String>>,
    pub date: Option<Option<String>>,
    pub highlight: Option<String>,
}

#[derive(Debug, Clone)]
pub struct MigrationWorkflowReport {
    pub xml_parts_checked: usize,
    pub published_output: PathBuf,
    pub text_exports_match: bool,
    pub word_fidelity_match: bool,
}

#[derive(Debug, Clone)]
pub struct FidelityReport {
    pub issues: Vec<String>,
}

impl FidelityReport {
    pub fn is_valid(&self) -> bool {
        self.issues.is_empty()
    }
}

#[derive(Debug, Clone, Deserialize, PartialEq, Eq, PartialOrd, Ord)]
struct WordParagraphSignature {
    text: String,
    style: Option<String>,
    highlights: Vec<String>,
}

impl AutomationSpec {
    pub fn from_path(path: &Path) -> Result<Self> {
        let raw = fs::read_to_string(path)
            .with_context(|| format!("failed to read spec file {}", path.display()))?;
        let spec: Self = serde_json::from_str(&raw)
            .with_context(|| format!("failed to parse JSON spec {}", path.display()))?;
        let spec_base_dir = path.parent().unwrap_or(Path::new("."));
        spec.validate(spec_base_dir)?;
        Ok(spec)
    }

    pub fn validate(&self, spec_base_dir: &Path) -> Result<()> {
        let mut issues = Vec::new();
        if self.operations.is_empty() {
            issues.push("spec must contain at least one operation".to_string());
        }
        for (index, operation) in self.operations.iter().enumerate() {
            validate_operation(operation, spec_base_dir, index, &mut issues);
        }
        if issues.is_empty() {
            Ok(())
        } else {
            bail!("invalid automation spec:\n{}", format_validation_issues(&issues))
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct WorkSessionMetadata {
    session_id: String,
    source_path: String,
    source_sha256: String,
    normalized_path: String,
    output_dir: String,
    current_version_path: String,
    detected_format: DocumentFormat,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(tag = "type", rename_all = "kebab-case")]
pub enum Operation {
    InsertParagraphs {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        entries: Vec<ParagraphEntry>,
    },
    ReplaceText {
        #[serde(default)]
        part: PartTarget,
        find: String,
        replace: String,
        #[serde(default = "default_highlight_value")]
        highlight: String,
    },
    DeleteParagraphs {
        #[serde(default)]
        part: PartTarget,
        contains: String,
    },
    InsertTableAfter {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        table: TableSpec,
    },
    InsertHyperlinkAfter {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        hyperlink: HyperlinkSpec,
    },
    InsertImageAfter {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        image: ImageSpec,
    },
    InsertSectionBreakAfter {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        break_type: SectionBreakType,
    },
    InsertCommentAfter {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        comment: CommentSpec,
    },
    InsertNoteAfter {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        kind: NoteKind,
        note: NoteSpec,
    },
    InsertContentControlAfter {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        control: ContentControlSpec,
    },
    InsertFieldAfter {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        field: FieldSpec,
    },
    TrackInsertParagraphs {
        #[serde(default)]
        part: PartTarget,
        anchor: AnchorTarget,
        author: String,
        date: String,
        entries: Vec<ParagraphEntry>,
    },
    UpsertHeaderFooter {
        kind: HeaderFooterKind,
        #[serde(default)]
        reference: HeaderFooterReferenceKind,
        #[serde(default)]
        section_index: usize,
        entries: Vec<ParagraphEntry>,
    },
    TrackDeleteParagraphs {
        #[serde(default)]
        part: PartTarget,
        contains: String,
        author: String,
        date: String,
    },
    SetCoreProperty {
        property: CoreProperty,
        value: String,
    },
}

#[derive(Debug, Clone, Default, Deserialize)]
pub struct PartTarget(#[serde(default)] pub Option<String>);

impl PartTarget {
    fn resolve(&self) -> Result<String> {
        resolve_part_name(self.0.as_deref())
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(untagged)]
pub enum AnchorTarget {
    Simple(String),
    Structured(AnchorSpec),
}

impl AnchorTarget {
    fn as_spec(&self) -> AnchorSpec {
        match self {
            AnchorTarget::Simple(text) => AnchorSpec {
                text: text.clone(),
                mode: AnchorMatchMode::Contains,
                occurrence: 1,
            },
            AnchorTarget::Structured(spec) => spec.clone(),
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
pub struct AnchorSpec {
    pub text: String,
    #[serde(default)]
    pub mode: AnchorMatchMode,
    #[serde(default = "default_occurrence")]
    pub occurrence: usize,
}

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum AnchorMatchMode {
    #[default]
    Contains,
    Equals,
    StartsWith,
    EndsWith,
}

fn default_occurrence() -> usize {
    1
}

#[derive(Debug, Clone, Deserialize)]
pub struct ParagraphEntry {
    pub text: String,
    #[serde(default)]
    pub style: ParagraphStyle,
    #[serde(default = "default_highlight_value")]
    pub highlight: String,
    #[serde(default)]
    pub bold: bool,
    #[serde(default)]
    pub italic: bool,
    #[serde(default)]
    pub underline: bool,
}

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum ParagraphStyle {
    #[default]
    Normal,
    Heading1,
    Heading2,
    Heading3,
    ListBullet,
    ListNumber,
    Quote,
}

impl ParagraphStyle {
    fn style_id(&self) -> &'static str {
        match self {
            ParagraphStyle::Normal => "Normal",
            ParagraphStyle::Heading1 => "Heading1",
            ParagraphStyle::Heading2 => "Heading2",
            ParagraphStyle::Heading3 => "Heading3",
            ParagraphStyle::ListBullet => "ListBullet",
            ParagraphStyle::ListNumber => "ListNumber",
            ParagraphStyle::Quote => "Quote",
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
pub struct TableSpec {
    pub rows: Vec<TableRowSpec>,
    #[serde(default)]
    pub style: Option<String>,
    #[serde(default = "default_highlight_value")]
    pub highlight: String,
}

#[derive(Debug, Clone, Deserialize)]
pub struct TableRowSpec {
    pub cells: Vec<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct HyperlinkSpec {
    pub text: String,
    pub url: String,
    #[serde(default)]
    pub style: ParagraphStyle,
    #[serde(default = "default_highlight_value")]
    pub highlight: String,
}

#[derive(Debug, Clone, Deserialize)]
pub struct ImageSpec {
    pub path: String,
    pub width_emu: u64,
    pub height_emu: u64,
    #[serde(default)]
    pub alt_text: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct CommentSpec {
    pub text: String,
    pub comment_text: String,
    pub author: String,
    #[serde(default)]
    pub initials: Option<String>,
    #[serde(default)]
    pub date: Option<String>,
    #[serde(default)]
    pub style: ParagraphStyle,
    #[serde(default = "default_highlight_value")]
    pub highlight: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NoteKind {
    Footnote,
    Endnote,
}

#[derive(Debug, Clone, Deserialize)]
pub struct NoteSpec {
    pub reference_text: String,
    pub body: String,
    #[serde(default)]
    pub style: ParagraphStyle,
    #[serde(default = "default_highlight_value")]
    pub highlight: String,
}

#[derive(Debug, Clone, Deserialize)]
pub struct ContentControlSpec {
    pub tag: String,
    #[serde(default)]
    pub alias: Option<String>,
    pub text: String,
    #[serde(default)]
    pub placeholder: Option<String>,
    #[serde(default)]
    pub style: ParagraphStyle,
    #[serde(default = "default_highlight_value")]
    pub highlight: String,
}

#[derive(Debug, Clone, Deserialize)]
pub struct FieldSpec {
    pub instruction: String,
    pub result: String,
    #[serde(default)]
    pub style: ParagraphStyle,
    #[serde(default = "default_highlight_value")]
    pub highlight: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum HeaderFooterKind {
    Header,
    Footer,
}

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum HeaderFooterReferenceKind {
    #[default]
    Default,
    First,
    Even,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum SectionBreakType {
    Continuous,
    NextPage,
    EvenPage,
    OddPage,
}

impl SectionBreakType {
    fn word_value(&self) -> &'static str {
        match self {
            SectionBreakType::Continuous => "continuous",
            SectionBreakType::NextPage => "nextPage",
            SectionBreakType::EvenPage => "evenPage",
            SectionBreakType::OddPage => "oddPage",
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum CoreProperty {
    Title,
    Subject,
    Creator,
    Description,
    Keywords,
    LastModifiedBy,
}

impl CoreProperty {
    fn xml_name(&self) -> &'static str {
        match self {
            CoreProperty::Title => "dc:title",
            CoreProperty::Subject => "dc:subject",
            CoreProperty::Creator => "dc:creator",
            CoreProperty::Description => "dc:description",
            CoreProperty::Keywords => "cp:keywords",
            CoreProperty::LastModifiedBy => "cp:lastModifiedBy",
        }
    }

    fn local_name(&self) -> &'static str {
        match self {
            CoreProperty::Title => "title",
            CoreProperty::Subject => "subject",
            CoreProperty::Creator => "creator",
            CoreProperty::Description => "description",
            CoreProperty::Keywords => "keywords",
            CoreProperty::LastModifiedBy => "lastModifiedBy",
        }
    }
}

fn default_highlight_value() -> String {
    DEFAULT_HIGHLIGHT.to_string()
}

fn validate_operation(
    operation: &Operation,
    spec_base_dir: &Path,
    index: usize,
    issues: &mut Vec<String>,
) {
    let prefix = format!("operation #{index}");
    match operation {
        Operation::InsertParagraphs {
            part,
            anchor,
            entries,
        } => {
            validate_part_target(part, &prefix, issues);
            validate_anchor_target(anchor, &prefix, issues);
            if entries.is_empty() {
                issues.push(format!("{prefix}: insert-paragraphs requires at least one entry"));
            }
            for (entry_index, entry) in entries.iter().enumerate() {
                validate_paragraph_entry(entry, &format!("{prefix} entry #{entry_index}"), issues);
            }
        }
        Operation::ReplaceText {
            part,
            find,
            highlight,
            ..
        } => {
            validate_part_target(part, &prefix, issues);
            validate_non_empty(find, &format!("{prefix}: replace-text find"), issues);
            validate_highlight(highlight, &format!("{prefix}: replace-text highlight"), issues);
        }
        Operation::DeleteParagraphs { part, contains }
        | Operation::TrackDeleteParagraphs { part, contains, .. } => {
            validate_part_target(part, &prefix, issues);
            validate_non_empty(contains, &format!("{prefix}: paragraph filter"), issues);
        }
        Operation::InsertTableAfter {
            part,
            anchor,
            table,
        } => {
            validate_part_target(part, &prefix, issues);
            validate_anchor_target(anchor, &prefix, issues);
            if table.rows.is_empty() {
                issues.push(format!("{prefix}: insert-table-after requires at least one row"));
            }
            for (row_index, row) in table.rows.iter().enumerate() {
                if row.cells.is_empty() {
                    issues.push(format!(
                        "{prefix}: table row #{row_index} requires at least one cell"
                    ));
                }
            }
            validate_highlight(&table.highlight, &format!("{prefix}: table highlight"), issues);
        }
        Operation::InsertHyperlinkAfter {
            part,
            anchor,
            hyperlink,
        } => {
            validate_part_target(part, &prefix, issues);
            validate_anchor_target(anchor, &prefix, issues);
            validate_non_empty(&hyperlink.text, &format!("{prefix}: hyperlink text"), issues);
            validate_non_empty(&hyperlink.url, &format!("{prefix}: hyperlink url"), issues);
            validate_highlight(
                &hyperlink.highlight,
                &format!("{prefix}: hyperlink highlight"),
                issues,
            );
        }
        Operation::InsertImageAfter {
            part,
            anchor,
            image,
        } => {
            validate_part_target(part, &prefix, issues);
            validate_anchor_target(anchor, &prefix, issues);
            validate_non_empty(&image.path, &format!("{prefix}: image path"), issues);
            let image_path = spec_base_dir.join(&image.path);
            if !image_path.exists() {
                issues.push(format!(
                    "{prefix}: image file does not exist: {}",
                    image_path.display()
                ));
            }
            if image_path.extension().and_then(|ext| ext.to_str()).is_none() {
                issues.push(format!(
                    "{prefix}: image path is missing a file extension: {}",
                    image_path.display()
                ));
            }
        }
        Operation::InsertSectionBreakAfter { part, anchor, .. }
        | Operation::InsertCommentAfter { part, anchor, .. }
        | Operation::InsertNoteAfter { part, anchor, .. }
        | Operation::InsertContentControlAfter { part, anchor, .. }
        | Operation::InsertFieldAfter { part, anchor, .. }
        | Operation::TrackInsertParagraphs { part, anchor, .. } => {
            validate_part_target(part, &prefix, issues);
            validate_anchor_target(anchor, &prefix, issues);
            match operation {
                Operation::InsertCommentAfter { comment, .. } => {
                    validate_non_empty(&comment.text, &format!("{prefix}: comment text"), issues);
                    validate_non_empty(
                        &comment.comment_text,
                        &format!("{prefix}: comment body"),
                        issues,
                    );
                    validate_non_empty(&comment.author, &format!("{prefix}: comment author"), issues);
                    validate_highlight(
                        &comment.highlight,
                        &format!("{prefix}: comment highlight"),
                        issues,
                    );
                }
                Operation::InsertNoteAfter { note, .. } => {
                    validate_non_empty(
                        &note.reference_text,
                        &format!("{prefix}: note reference text"),
                        issues,
                    );
                    validate_non_empty(&note.body, &format!("{prefix}: note body"), issues);
                    validate_highlight(
                        &note.highlight,
                        &format!("{prefix}: note highlight"),
                        issues,
                    );
                }
                Operation::InsertContentControlAfter { control, .. } => {
                    validate_non_empty(&control.tag, &format!("{prefix}: content control tag"), issues);
                    validate_non_empty(
                        &control.text,
                        &format!("{prefix}: content control text"),
                        issues,
                    );
                    validate_highlight(
                        &control.highlight,
                        &format!("{prefix}: content control highlight"),
                        issues,
                    );
                }
                Operation::InsertFieldAfter { field, .. } => {
                    validate_non_empty(
                        &field.instruction,
                        &format!("{prefix}: field instruction"),
                        issues,
                    );
                    validate_highlight(
                        &field.highlight,
                        &format!("{prefix}: field highlight"),
                        issues,
                    );
                }
                Operation::TrackInsertParagraphs {
                    author,
                    date,
                    entries,
                    ..
                } => {
                    validate_non_empty(author, &format!("{prefix}: tracked insert author"), issues);
                    validate_non_empty(date, &format!("{prefix}: tracked insert date"), issues);
                    if entries.is_empty() {
                        issues.push(format!(
                            "{prefix}: track-insert-paragraphs requires at least one entry"
                        ));
                    }
                    for (entry_index, entry) in entries.iter().enumerate() {
                        validate_paragraph_entry(
                            entry,
                            &format!("{prefix} tracked entry #{entry_index}"),
                            issues,
                        );
                    }
                }
                _ => {}
            }
        }
        Operation::UpsertHeaderFooter { entries, .. } => {
            if entries.is_empty() {
                issues.push(format!("{prefix}: upsert-header-footer requires at least one entry"));
            }
            for (entry_index, entry) in entries.iter().enumerate() {
                validate_paragraph_entry(entry, &format!("{prefix} entry #{entry_index}"), issues);
            }
        }
        Operation::SetCoreProperty { .. } => {}
    }

    if let Operation::TrackDeleteParagraphs { author, date, .. } = operation {
        validate_non_empty(author, &format!("{prefix}: tracked delete author"), issues);
        validate_non_empty(date, &format!("{prefix}: tracked delete date"), issues);
    }
}

fn validate_paragraph_entry(entry: &ParagraphEntry, label: &str, issues: &mut Vec<String>) {
    validate_non_empty(&entry.text, &format!("{label}: text"), issues);
    validate_highlight(&entry.highlight, &format!("{label}: highlight"), issues);
}

fn validate_anchor_target(anchor: &AnchorTarget, label: &str, issues: &mut Vec<String>) {
    let anchor = anchor.as_spec();
    validate_non_empty(&anchor.text, &format!("{label}: anchor text"), issues);
    if anchor.occurrence == 0 {
        issues.push(format!("{label}: anchor occurrence must be at least 1"));
    }
}

fn validate_part_target(part: &PartTarget, label: &str, issues: &mut Vec<String>) {
    if let Err(err) = part.resolve() {
        issues.push(format!("{label}: {err}"));
    }
}

fn validate_non_empty(value: &str, label: &str, issues: &mut Vec<String>) {
    if value.trim().is_empty() {
        issues.push(format!("{label} must not be empty"));
    }
}

fn validate_highlight(value: &str, label: &str, issues: &mut Vec<String>) {
    if !SUPPORTED_HIGHLIGHTS.contains(&value) {
        issues.push(format!(
            "{label} uses unsupported highlight '{value}'. Supported values: {}",
            SUPPORTED_HIGHLIGHTS.join(", ")
        ));
    }
}

fn sha256_hex(input: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(input);
    format!("{:x}", hasher.finalize())
}

fn ensure_session_id_is_safe(session_id: &str) -> Result<()> {
    let valid = !session_id.trim().is_empty()
        && session_id
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '-' | '_' | '.'));
    if valid {
        Ok(())
    } else {
        bail!("session id must contain only letters, numbers, '.', '-', or '_'");
    }
}

fn session_root(temp_root: &Path) -> PathBuf {
    temp_root.join("sessions")
}

fn session_dir(temp_root: &Path, session_id: &str) -> PathBuf {
    session_root(temp_root).join(session_id)
}

fn session_metadata_path(temp_root: &Path, session_id: &str) -> PathBuf {
    session_dir(temp_root, session_id).join(SESSION_METADATA_FILE)
}

fn load_work_session_metadata(temp_root: &Path, session_id: &str) -> Result<WorkSessionMetadata> {
    let metadata_path = session_metadata_path(temp_root, session_id);
    let raw = fs::read_to_string(&metadata_path)
        .with_context(|| format!("failed to read session metadata {}", metadata_path.display()))?;
    serde_json::from_str(&raw)
        .with_context(|| format!("failed to parse session metadata {}", metadata_path.display()))
}

fn save_work_session_metadata(temp_root: &Path, metadata: &WorkSessionMetadata) -> Result<()> {
    let metadata_path = session_metadata_path(temp_root, &metadata.session_id);
    if let Some(parent) = metadata_path.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("failed to create session dir {}", parent.display()))?;
    }
    let content = serde_json::to_string_pretty(metadata).context("failed to serialize session metadata")?;
    fs::write(&metadata_path, content)
        .with_context(|| format!("failed to write session metadata {}", metadata_path.display()))
}

fn parse_versioned_stem(path: &Path) -> Result<(String, usize, usize)> {
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .with_context(|| format!("path must have a UTF-8 file stem: {}", path.display()))?;
    let re = Regex::new(r"^(?P<prefix>.*?)(?P<version>v(?P<number>\d+))$")
        .context("invalid version regex")?;
    let captures = re
        .captures(stem)
        .with_context(|| format!("path does not end with a version marker like v001: {}", path.display()))?;
    let prefix = captures
        .name("prefix")
        .map(|value| value.as_str().to_string())
        .unwrap_or_default();
    let number = captures
        .name("number")
        .context("missing version number capture")?
        .as_str()
        .parse::<usize>()
        .context("failed to parse version number")?;
    let width = captures
        .name("number")
        .context("missing version width capture")?
        .as_str()
        .len();
    Ok((prefix, number, width))
}

fn highest_versioned_output_path(current_path: &Path, output_dir: Option<&Path>) -> Result<(PathBuf, usize, usize)> {
    let (prefix, current_version, width) = parse_versioned_stem(current_path)?;
    let extension = current_path
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_string())
        .with_context(|| format!("path must include an extension: {}", current_path.display()))?;
    let directory = output_dir
        .map(Path::to_path_buf)
        .or_else(|| current_path.parent().map(Path::to_path_buf))
        .context("unable to determine output directory for publish-next")?;
    fs::create_dir_all(&directory)
        .with_context(|| format!("failed to create output directory {}", directory.display()))?;

    let pattern = Regex::new(&format!(
        r"^{}v(?P<number>\d+)\.{}$",
        regex::escape(&prefix),
        regex::escape(&extension)
    ))
    .context("invalid next-version regex")?;

    let mut max_version = current_version;
    for entry in fs::read_dir(&directory)
        .with_context(|| format!("failed to read output directory {}", directory.display()))?
    {
        let entry = entry.with_context(|| format!("failed to inspect {}", directory.display()))?;
        let file_name = entry.file_name();
        let Some(file_name) = file_name.to_str() else {
            continue;
        };
        if let Some(captures) = pattern.captures(file_name)
            && let Some(number) = captures.name("number")
            && let Ok(version) = number.as_str().parse::<usize>()
        {
            max_version = max_version.max(version);
        }
    }

    let file_name = format!("{prefix}v{:0width$}.{extension}", max_version, width = width);
    Ok((directory.join(file_name), max_version, width))
}

fn resolve_versioned_output_path(
    current_path: &Path,
    output_dir: Option<&Path>,
    mode: PublishTargetMode,
) -> Result<(PathBuf, usize)> {
    let (highest_path, highest_version, width) = highest_versioned_output_path(current_path, output_dir)?;
    let (prefix, _, _) = parse_versioned_stem(current_path)?;
    let extension = current_path
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_string())
        .with_context(|| format!("path must include an extension: {}", current_path.display()))?;
    let directory = highest_path
        .parent()
        .map(Path::to_path_buf)
        .context("unable to determine resolved output directory")?;

    match mode {
        PublishTargetMode::Latest => Ok((highest_path, highest_version)),
        PublishTargetMode::NextVersion => {
            let next_version = highest_version + 1;
            let file_name = format!("{prefix}v{:0width$}.{extension}", next_version, width = width);
            Ok((directory.join(file_name), next_version))
        }
    }
}

pub fn inspect_normalization_file(input: &Path) -> Result<NormalizationReport> {
    let bytes = fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    Ok(inspect_normalization_bytes(&bytes))
}

fn verify_published_output_matches_candidate(output: &Path, candidate_bytes: &[u8]) -> Result<()> {
    let published_bytes =
        fs::read(output).with_context(|| format!("failed to read published output {}", output.display()))?;
    let normalization = inspect_normalization_bytes(&published_bytes);
    if !normalization.is_normalized {
        bail!(
            "published output mutated after write and is no longer automation-safe ({}; format={}):\n{}",
            output.display(),
            normalization.format.as_str(),
            format_validation_issues(&normalization.details)
        );
    }
    if published_bytes != candidate_bytes {
        bail!(
            "published output does not match the validated candidate after write: {}",
            output.display()
        );
    }
    Ok(())
}

fn ensure_existing_output_is_safe_for_direct_write(output: &Path) -> Result<()> {
    if !output.exists() {
        return Ok(());
    }

    let normalization = inspect_normalization_file(output)?;
    if normalization.is_normalized {
        return Ok(());
    }

    bail!(
        "refusing to overwrite existing non-normalized or protection-sensitive output in place ({}; format={}):\n{}",
        output.display(),
        normalization.format.as_str(),
        format_validation_issues(&normalization.details)
    );
}

pub fn normalize_docx_file(
    input: &Path,
    output: &Path,
    temp_root: Option<&Path>,
    trusted_ooxml: Option<&Path>,
    allow_word_com_encrypted_package: bool,
) -> Result<NormalizeWorkflowReport> {
    let bytes = fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    let normalization = inspect_normalization_bytes(&bytes);

    match normalization.format {
        DocumentFormat::OoxmlZip => {
            let report = validate_docx_bytes(&bytes)?;
            if !report.is_valid() {
                bail!(
                    "normalized source failed validation:\n{}",
                    format_validation_issues(&report.issues)
                );
            }
            if let Some(parent) = output.parent()
                && !parent.as_os_str().is_empty()
            {
                fs::create_dir_all(parent).with_context(|| {
                    format!("failed to create output directory {}", parent.display())
                })?;
            }
            fs::copy(input, output).with_context(|| {
                format!(
                    "failed to publish normalized output {} -> {}",
                    input.display(),
                    output.display()
                )
            })?;
            verify_published_output_matches_candidate(output, &bytes)?;
            Ok(NormalizeWorkflowReport {
                detected_format: normalization.format,
                already_normalized: true,
                xml_parts_checked: report.xml_parts_checked,
                published_output: output.to_path_buf(),
            })
        }
        DocumentFormat::OleEncryptedPackage if allow_word_com_encrypted_package => {
            let report = migrate_source_to_ooxml(input, output, temp_root, trusted_ooxml)?;
            Ok(NormalizeWorkflowReport {
                detected_format: normalization.format,
                already_normalized: false,
                xml_parts_checked: report.xml_parts_checked,
                published_output: report.published_output,
            })
        }
        _ => {
            let _ = temp_root;
            let _ = trusted_ooxml;
            bail!(
                "source document requires normalization but this build does not yet include a reliable non-COM normalization backend for format '{}'. For EncryptedPackage sources, rerun `normalize` with the explicit COM exception flag if you want to allow the guarded fallback:\n{}",
                normalization.format.as_str(),
                format_validation_issues(&normalization.details)
            );
        }
    }
}

pub fn validate_spec_file(spec_path: &Path) -> Result<usize> {
    let spec = AutomationSpec::from_path(spec_path)?;
    Ok(spec.operations.len())
}

pub fn add_comment_to_docx(
    input: &Path,
    output: &Path,
    part: Option<&str>,
    anchor: &AnchorTarget,
    comment: &CommentSpec,
) -> Result<()> {
    let spec = AutomationSpec {
        operations: vec![Operation::InsertCommentAfter {
            part: PartTarget(part.map(|value| value.to_string())),
            anchor: anchor.clone(),
            comment: comment.clone(),
        }],
    };
    apply_spec_file_to_docx(input, output, &spec, Path::new("."))
}

pub fn list_docx_comments(input: &Path) -> Result<Vec<CommentRecord>> {
    let input_bytes =
        fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    ensure_input_is_normalized(input, &input_bytes)?;
    let package = DocxPackage::read(&input_bytes)?;
    read_comments_from_package(&package)
}

pub fn update_docx_comment(
    input: &Path,
    output: &Path,
    comment_id: u32,
    update: &CommentUpdate,
) -> Result<()> {
    let input_bytes =
        fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    ensure_input_is_normalized(input, &input_bytes)?;
    let mut package = DocxPackage::read(&input_bytes)?;
    let mut updated = false;
    with_xml_part_mut(&mut package, "word/comments.xml", |root| {
        for child in &mut root.children {
            let XMLNode::Element(element) = child else {
                continue;
            };
            if local_name(&element.name) != "comment" || comment_id_of_element(element) != Some(comment_id)
            {
                continue;
            }
            let mut record = comment_record_from_element(element);
            if let Some(value) = &update.comment_text {
                record.comment_text = value.clone();
            }
            if let Some(value) = &update.author {
                record.author = value.clone();
            }
            if let Some(value) = &update.initials {
                record.initials = value.clone();
            }
            if let Some(value) = &update.date {
                record.date = value.clone();
            }
            if let Some(value) = &update.highlight {
                record.highlight = value.clone();
            }
            let mut issues = Vec::new();
            validate_non_empty(&record.comment_text, "updated comment text", &mut issues);
            validate_highlight(&record.highlight, "updated comment highlight", &mut issues);
            if !issues.is_empty() {
                bail!("{}", format_validation_issues(&issues));
            }
            *element = make_comment_entry_from_record(&record)?;
            updated = true;
            break;
        }
        Ok(())
    })?;
    if !updated {
        bail!("comment not found: {comment_id}");
    }
    write_updated_package(output, package)
}

pub fn delete_docx_comment(input: &Path, output: &Path, comment_id: u32) -> Result<()> {
    let input_bytes =
        fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    ensure_input_is_normalized(input, &input_bytes)?;
    let mut package = DocxPackage::read(&input_bytes)?;
    let mut removed = false;

    if package.get_file("word/comments.xml").is_some() {
        with_xml_part_mut(&mut package, "word/comments.xml", |root| {
            let before = root.children.len();
            root.children.retain(|child| {
                !matches!(child, XMLNode::Element(element)
                    if local_name(&element.name) == "comment"
                    && comment_id_of_element(element) == Some(comment_id))
            });
            removed = root.children.len() != before;
            Ok(())
        })?;
    }

    if !removed {
        bail!("comment not found: {comment_id}");
    }

    for part_name in comment_story_part_names(&package) {
        with_xml_part_mut(&mut package, &part_name, |root| {
            remove_comment_markup_from_element(root, comment_id);
            Ok(())
        })?;
    }

    write_updated_package(output, package)
}

fn build_validated_candidate_from_source_bytes(
    source_path: &Path,
    source_bytes: &[u8],
    spec: &AutomationSpec,
    spec_base_dir: &Path,
) -> Result<(Vec<u8>, ValidationReport)> {
    spec.validate(spec_base_dir)?;
    ensure_input_is_normalized(source_path, source_bytes)?;
    let source_validation = validate_docx_bytes(source_bytes)?;
    if !source_validation.is_valid() {
        bail!(
            "source document is not automation-safe:\n{}",
            format_validation_issues(&source_validation.issues)
        );
    }

    let output_bytes = apply_spec_to_docx_bytes(source_bytes, spec, spec_base_dir)?;
    let candidate_validation = validate_docx_bytes(&output_bytes)?;
    if !candidate_validation.is_valid() {
        bail!(
            "candidate output failed validation:\n{}",
            format_validation_issues(&candidate_validation.issues)
        );
    }

    let fidelity_report = validate_source_fidelity_bytes(source_bytes, &output_bytes, spec)?;
    if !fidelity_report.is_valid() {
        bail!(
            "candidate output failed source fidelity checks:\n{}",
            format_validation_issues(&fidelity_report.issues)
        );
    }

    Ok((output_bytes, candidate_validation))
}

fn write_updated_package(output: &Path, package: DocxPackage) -> Result<()> {
    let output_bytes = package.write()?;
    let validation = validate_docx_bytes(&output_bytes)?;
    if !validation.is_valid() {
        bail!(
            "updated output failed validation:\n{}",
            format_validation_issues(&validation.issues)
        );
    }
    fs::write(output, output_bytes).with_context(|| format!("failed to write {}", output.display()))
}

pub fn apply_spec_file_to_docx(
    input: &Path,
    output: &Path,
    spec: &AutomationSpec,
    spec_base_dir: &Path,
) -> Result<()> {
    let input_bytes =
        fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    spec.validate(spec_base_dir)?;
    ensure_input_is_normalized(input, &input_bytes)?;
    let output_bytes = apply_spec_to_docx_bytes(&input_bytes, spec, spec_base_dir)?;
    fs::write(output, output_bytes)
        .with_context(|| format!("failed to write {}", output.display()))?;
    Ok(())
}

fn publish_spec_file_to_docx_internal(
    input: &Path,
    output: &Path,
    spec: &AutomationSpec,
    spec_base_dir: &Path,
    temp_root: Option<&Path>,
    allow_update_latest: bool,
) -> Result<PublishWorkflowReport> {
    if input == output && !allow_update_latest {
        bail!(
            "input and output must be different paths for publish workflow: {}",
            input.display()
        );
    }

    let temp_dir = create_work_dir(temp_root.unwrap_or(Path::new(DEFAULT_TEMP_ROOT)))?;
    let staged_input = temp_dir.join("source.docx");
    let candidate_output = temp_dir.join("candidate.docx");

    let workflow = (|| -> Result<PublishWorkflowReport> {
        spec.validate(spec_base_dir)?;
        fs::copy(input, &staged_input).with_context(|| {
            format!(
                "failed to stage source into temporary workspace: {} -> {}",
                input.display(),
                staged_input.display()
            )
        })?;

        let source_bytes = fs::read(&staged_input)
            .with_context(|| format!("failed to read staged source {}", staged_input.display()))?;
        let (output_bytes, candidate_validation) = build_validated_candidate_from_source_bytes(
            &staged_input,
            &source_bytes,
            spec,
            spec_base_dir,
        )?;
        fs::write(&candidate_output, &output_bytes).with_context(|| {
            format!(
                "failed to write temporary candidate output {}",
                candidate_output.display()
            )
        })?;

        if let Some(parent) = output.parent()
            && !parent.as_os_str().is_empty()
        {
            fs::create_dir_all(parent).with_context(|| {
                format!("failed to create output directory {}", parent.display())
            })?;
        }
        ensure_existing_output_is_safe_for_direct_write(output)?;

        fs::copy(&candidate_output, output).with_context(|| {
            format!(
                "failed to publish validated output {} -> {}",
                candidate_output.display(),
                output.display()
            )
        })?;
        verify_published_output_matches_candidate(output, &output_bytes)?;

        Ok(PublishWorkflowReport {
            xml_parts_checked: candidate_validation.xml_parts_checked,
            published_output: output.to_path_buf(),
        })
    })();

    match workflow {
        Ok(report) => {
            cleanup_temp_dir(&temp_dir)?;
            Ok(report)
        }
        Err(err) => Err(err.context(format!(
            "publish workflow aborted; temporary workspace preserved at {}",
            temp_dir.display()
        ))),
    }
}

pub fn publish_spec_file_to_docx(
    input: &Path,
    output: &Path,
    spec: &AutomationSpec,
    spec_base_dir: &Path,
    temp_root: Option<&Path>,
) -> Result<PublishWorkflowReport> {
    publish_spec_file_to_docx_internal(input, output, spec, spec_base_dir, temp_root, false)
}

pub fn migrate_source_to_ooxml(
    input: &Path,
    output: &Path,
    temp_root: Option<&Path>,
    trusted_ooxml: Option<&Path>,
) -> Result<MigrationWorkflowReport> {
    if input == output {
        bail!(
            "input and output must be different paths for migration workflow: {}",
            input.display()
        );
    }

    let temp_dir = create_work_dir(temp_root.unwrap_or(Path::new(r"C:\Temp\wordflow")))?;
    let staged_input = temp_dir.join("source-input.docx");
    let source_export = temp_dir.join("source-export.txt");
    let source_signatures = temp_dir.join("source-signatures.json");
    let rtf_output = temp_dir.join("migration-stage.rtf");
    let migrated_output = temp_dir.join("migrated.docx");
    let migrated_export = temp_dir.join("migrated-export.txt");
    let migrated_signatures = temp_dir.join("migrated-signatures.json");
    let script_path = temp_dir.join("run-migration.ps1");

    let workflow = (|| -> Result<MigrationWorkflowReport> {
        fs::copy(input, &staged_input).with_context(|| {
            format!(
                "failed to stage source into temporary workspace: {} -> {}",
                input.display(),
                staged_input.display()
            )
        })?;

        write_word_migration_script(
            &script_path,
            &staged_input,
            &source_export,
            &source_signatures,
            &rtf_output,
            &migrated_output,
            &migrated_export,
            &migrated_signatures,
        )?;
        run_powershell_file(&script_path)?;

        let source_text = fs::read(&source_export).with_context(|| {
            format!("failed to read source text export {}", source_export.display())
        })?;
        let migrated_text = fs::read(&migrated_export).with_context(|| {
            format!(
                "failed to read migrated text export {}",
                migrated_export.display()
            )
        })?;
        if source_text != migrated_text {
            bail!(
                "migration text export mismatch between {} and {}",
                source_export.display(),
                migrated_export.display()
            );
        }

        let word_fidelity = validate_word_fidelity_exports(
            &source_signatures,
            &migrated_signatures,
        )?;
        if !word_fidelity.is_valid() {
            bail!(
                "migrated candidate failed Word fidelity checks:\n{}",
                format_validation_issues(&word_fidelity.issues)
            );
        }

        let candidate_validation = validate_docx_file(&migrated_output)?;
        if !candidate_validation.is_valid() {
            bail!(
                "migrated candidate failed validation:\n{}",
                format_validation_issues(&candidate_validation.issues)
            );
        }

        if let Some(reference) = trusted_ooxml {
            let fidelity = validate_source_fidelity_file(reference, &migrated_output, None)?;
            if !fidelity.is_valid() {
                bail!(
                    "migrated candidate failed source fidelity checks:\n{}",
                    format_validation_issues(&fidelity.issues)
                );
            }
        }

        if let Some(parent) = output.parent()
            && !parent.as_os_str().is_empty()
        {
            fs::create_dir_all(parent).with_context(|| {
                format!("failed to create output directory {}", parent.display())
            })?;
        }
        ensure_existing_output_is_safe_for_direct_write(output)?;

        fs::copy(&migrated_output, output).with_context(|| {
            format!(
                "failed to publish migrated output {} -> {}",
                migrated_output.display(),
                output.display()
            )
        })?;
        let migrated_bytes = fs::read(&migrated_output).with_context(|| {
            format!("failed to read validated migrated candidate {}", migrated_output.display())
        })?;
        verify_published_output_matches_candidate(output, &migrated_bytes)?;

        Ok(MigrationWorkflowReport {
            xml_parts_checked: candidate_validation.xml_parts_checked,
            published_output: output.to_path_buf(),
            text_exports_match: true,
            word_fidelity_match: true,
        })
    })();

    match workflow {
        Ok(report) => {
            cleanup_temp_dir(&temp_dir)?;
            Ok(report)
        }
        Err(err) => Err(err.context(format!(
            "migration workflow aborted; temporary workspace preserved at {}",
            temp_dir.display()
        ))),
    }
}

pub fn apply_spec_file_to_docx_dry_run(
    input: &Path,
    spec: &AutomationSpec,
    spec_base_dir: &Path,
) -> Result<ValidationReport> {
    let input_bytes =
        fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    ensure_input_is_normalized(input, &input_bytes)?;
    let output_bytes = apply_spec_to_docx_bytes(&input_bytes, spec, spec_base_dir)?;
    validate_docx_bytes(&output_bytes)
}

pub fn list_docx_parts(input: &Path) -> Result<Vec<PartSummary>> {
    let bytes = fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    ensure_input_is_normalized(input, &bytes)?;
    let package = DocxPackage::read(&bytes)?;
    Ok(package
        .entries
        .iter()
        .map(|entry| PartSummary {
            name: entry.name.clone(),
            is_dir: entry.is_dir,
            size: entry.bytes.len(),
            sha256: (!entry.is_dir).then(|| {
                let mut hasher = Sha256::new();
                hasher.update(&entry.bytes);
                format!("{:x}", hasher.finalize())
            }),
        })
        .collect())
}

pub fn find_anchors_in_docx(
    input: &Path,
    part_filter: Option<&str>,
    anchor: &AnchorTarget,
) -> Result<Vec<AnchorMatch>> {
    let bytes = fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    ensure_input_is_normalized(input, &bytes)?;
    let package = DocxPackage::read(&bytes)?;
    let spec = anchor.as_spec();
    let mut matches = Vec::new();
    for entry in &package.entries {
        if entry.is_dir || !entry.name.ends_with(".xml") {
            continue;
        }
        if let Some(filter) = part_filter
            && entry.name != filter
        {
            continue;
        }
        if let Ok(root) = parse_xml(&entry.bytes)
            && let Ok(story) = story_container_ref(&root)
        {
            let mut matched_count = 0usize;
            for (idx, node) in story.children.iter().enumerate() {
                if let XMLNode::Element(element) = node
                    && local_name(&element.name) == "p"
                {
                    let text = paragraph_text(element);
                    if anchor_matches(&text, &spec) {
                        matched_count += 1;
                        if matched_count == spec.occurrence {
                            matches.push(AnchorMatch {
                                part: entry.name.clone(),
                                index: idx,
                                text,
                            });
                        }
                    }
                }
            }
        }
    }
    Ok(matches)
}

pub fn validate_docx_file(input: &Path) -> Result<ValidationReport> {
    let bytes = fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    ensure_input_is_normalized(input, &bytes)?;
    validate_docx_bytes(&bytes)
}

pub fn validate_source_fidelity_bytes(
    source_bytes: &[u8],
    candidate_bytes: &[u8],
    spec: &AutomationSpec,
) -> Result<FidelityReport> {
    validate_source_fidelity_bytes_with_optional_spec(source_bytes, candidate_bytes, Some(spec))
}

pub fn validate_source_fidelity_file(
    source: &Path,
    candidate: &Path,
    spec: Option<&AutomationSpec>,
) -> Result<FidelityReport> {
    let source_bytes =
        fs::read(source).with_context(|| format!("failed to read {}", source.display()))?;
    let candidate_bytes =
        fs::read(candidate).with_context(|| format!("failed to read {}", candidate.display()))?;
    validate_source_fidelity_bytes_with_optional_spec(&source_bytes, &candidate_bytes, spec)
}

fn validate_source_fidelity_bytes_with_optional_spec(
    source_bytes: &[u8],
    candidate_bytes: &[u8],
    spec: Option<&AutomationSpec>,
) -> Result<FidelityReport> {
    let source_package = DocxPackage::read(source_bytes)?;
    let candidate_package = DocxPackage::read(candidate_bytes)?;
    let mut issues = Vec::new();

    for source_entry in &source_package.entries {
        if source_entry.is_dir || !is_story_part(&source_entry.name) {
            continue;
        }

        let Some(candidate_bytes) = candidate_package.get_file(&source_entry.name) else {
            issues.push(format!("missing story part in candidate: {}", source_entry.name));
            continue;
        };

        let source_root = parse_xml(&source_entry.bytes)?;
        let candidate_root = parse_xml(candidate_bytes)?;
        let source_story = story_container_ref(&source_root)?;
        let candidate_story = story_container_ref(&candidate_root)?;
        let source_signatures = paragraph_signatures(
            source_story,
            |paragraph| {
                spec.is_some_and(|spec| {
                    should_skip_paragraph_fidelity_check(&source_entry.name, paragraph, spec)
                })
            },
        );
        let candidate_signatures = paragraph_signatures(candidate_story, |_| false);

        for (signature, source_count) in source_signatures {
            let candidate_count = candidate_signatures.get(&signature).copied().unwrap_or(0);
            if candidate_count < source_count {
                issues.push(format!(
                    "formatting changed or disappeared in {} for paragraph '{}' (expected at least {}, found {})",
                    source_entry.name, signature.text, source_count, candidate_count
                ));
            }
        }
    }

    Ok(FidelityReport { issues })
}

fn ensure_input_is_normalized(input: &Path, input_bytes: &[u8]) -> Result<()> {
    let normalization = inspect_normalization_bytes(input_bytes);
    if normalization.is_normalized {
        return Ok(());
    }

    bail!(
        "source document is not normalized for OpenXML automation ({}):\n{}\nRun `normalize` first once a reliable non-COM normalization backend is available.",
        input.display(),
        format_validation_issues(&normalization.details)
    );
}

fn inspect_normalization_bytes(input_bytes: &[u8]) -> NormalizationReport {
    let format = detect_document_format(input_bytes);
    let details = match format {
        DocumentFormat::OoxmlZip => vec!["source is already a zip-based OOXML document".to_string()],
        DocumentFormat::OleEncryptedPackage => vec![
            "source is an OLE compound file with an EncryptedPackage payload".to_string(),
            "the document is not directly automation-safe for OpenXML editing".to_string(),
        ],
        DocumentFormat::OleWordBinary => vec![
            "source is an OLE Word binary document".to_string(),
            "the document is not directly automation-safe for OpenXML editing".to_string(),
        ],
        DocumentFormat::OleCompound => vec![
            "source is an OLE compound document".to_string(),
            "the document is not directly automation-safe for OpenXML editing".to_string(),
        ],
        DocumentFormat::Unknown => vec![
            "source format could not be recognized as OOXML or supported legacy Word packaging"
                .to_string(),
        ],
    };

    let is_normalized = format == DocumentFormat::OoxmlZip;
    NormalizationReport {
        format,
        is_normalized,
        requires_normalization: !is_normalized,
        details,
    }
}

fn detect_document_format(input_bytes: &[u8]) -> DocumentFormat {
    if input_bytes.starts_with(&ZIP_HEADER) {
        return DocumentFormat::OoxmlZip;
    }
    if input_bytes.starts_with(&OLE_HEADER) {
        if contains_utf16le_label(input_bytes, "EncryptedPackage") {
            return DocumentFormat::OleEncryptedPackage;
        }
        if contains_utf16le_label(input_bytes, "WordDocument") {
            return DocumentFormat::OleWordBinary;
        }
        return DocumentFormat::OleCompound;
    }
    DocumentFormat::Unknown
}

fn contains_utf16le_label(input_bytes: &[u8], text: &str) -> bool {
    let mut pattern = Vec::with_capacity(text.len() * 2);
    for unit in text.encode_utf16() {
        pattern.extend_from_slice(&unit.to_le_bytes());
    }
    input_bytes
        .windows(pattern.len())
        .any(|window| window == pattern.as_slice())
}

pub fn validate_docx_bytes(input_bytes: &[u8]) -> Result<ValidationReport> {
    let package = DocxPackage::read(input_bytes)?;
    let mut issues = Vec::new();
    let mut xml_parts_checked = 0usize;

    if package.get_file(DOCUMENT_XML_PATH).is_none() {
        issues.push(format!("missing {}", DOCUMENT_XML_PATH));
    }
    if package.get_file(CONTENT_TYPES_XML_PATH).is_none() {
        issues.push(format!("missing {}", CONTENT_TYPES_XML_PATH));
    }

    for entry in &package.entries {
        if entry.is_dir || !entry.name.ends_with(".xml") {
            continue;
        }
        xml_parts_checked += 1;
        match parse_xml(&entry.bytes) {
            Ok(root) => {
                if entry.name == DOCUMENT_XML_PATH && story_container_ref(&root).is_err() {
                    issues.push("word/document.xml missing body".to_string());
                }
            }
            Err(err) => issues.push(format!("{}: {}", entry.name, err)),
        }
    }

    Ok(ValidationReport {
        xml_parts_checked,
        issues,
    })
}

pub fn diff_docx_files(before: &Path, after: &Path) -> Result<DiffSummary> {
    let before_bytes =
        fs::read(before).with_context(|| format!("failed to read {}", before.display()))?;
    let after_bytes =
        fs::read(after).with_context(|| format!("failed to read {}", after.display()))?;
    ensure_input_is_normalized(before, &before_bytes)?;
    ensure_input_is_normalized(after, &after_bytes)?;
    let before_package = DocxPackage::read(&before_bytes)?;
    let after_package = DocxPackage::read(&after_bytes)?;

    let before_map = before_package.file_map();
    let after_map = after_package.file_map();

    let mut added_parts = Vec::new();
    let mut removed_parts = Vec::new();
    let mut changed_parts = Vec::new();

    for name in after_map.keys() {
        if !before_map.contains_key(name) {
            added_parts.push(name.clone());
        } else if before_map.get(name) != after_map.get(name) {
            changed_parts.push(name.clone());
        }
    }
    for name in before_map.keys() {
        if !after_map.contains_key(name) {
            removed_parts.push(name.clone());
        }
    }

    added_parts.sort();
    removed_parts.sort();
    changed_parts.sort();

    Ok(DiffSummary {
        added_parts,
        removed_parts,
        changed_parts,
    })
}

fn create_work_dir(temp_root: &Path) -> Result<PathBuf> {
    fs::create_dir_all(temp_root)
        .with_context(|| format!("failed to create temp root {}", temp_root.display()))?;

    let pid = process::id();
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis();
    for attempt in 0..100u32 {
        let candidate = temp_root.join(format!("run-{pid}-{stamp}-{attempt}"));
        if candidate.exists() {
            continue;
        }
        fs::create_dir(&candidate)
            .with_context(|| format!("failed to create temp dir {}", candidate.display()))?;
        return Ok(candidate);
    }

    bail!(
        "failed to allocate a unique temp workspace under {}",
        temp_root.display()
    )
}

fn cleanup_temp_dir(temp_dir: &Path) -> Result<()> {
    if temp_dir.exists() {
        fs::remove_dir_all(temp_dir)
            .with_context(|| format!("failed to clean temp dir {}", temp_dir.display()))?;
    }
    Ok(())
}

fn write_word_migration_script(
    script_path: &Path,
    staged_input: &Path,
    source_export: &Path,
    source_signatures: &Path,
    rtf_output: &Path,
    migrated_output: &Path,
    migrated_export: &Path,
    migrated_signatures: &Path,
) -> Result<()> {
    let script = format!(
        r#"$ErrorActionPreference = 'Stop'
function Get-ParagraphSignatureObjects($document) {{
  $items = @()
  foreach ($paragraph in $document.Paragraphs) {{
    $rawText = $paragraph.Range.Text
    if ($null -eq $rawText) {{ continue }}
    $text = $rawText -replace "[\r\a]+", "" -replace "\s+", " "
    $text = $text.Trim()
    if (-not $text) {{ continue }}

    $styleName = $null
    try {{
      if ($null -ne $paragraph.Range.Style) {{
        $styleName = [string]$paragraph.Range.Style.NameLocal
      }}
    }} catch {{
      $styleName = $null
    }}

    $highlights = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($character in $paragraph.Range.Characters) {{
      try {{
        $idx = [int]$character.HighlightColorIndex
        if ($idx -gt 0) {{
          [void]$highlights.Add($idx.ToString())
        }}
      }} catch {{
      }}
    }}

    $items += [pscustomobject]@{{
      text = $text
      style = $styleName
      highlights = @($highlights | Sort-Object)
    }}
  }}
  return $items
}}

$sourceSignatures = $null
$migratedSignatures = $null
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {{
  $doc = $word.Documents.Open({staged_input}, $false, $true)
  $doc.SaveAs2({source_export}, 2)
  $sourceSignatures = Get-ParagraphSignatureObjects $doc
  $doc.Close()

  $doc = $word.Documents.Open({staged_input}, $false, $true)
  $doc.SaveAs2({rtf_output}, 6)
  $doc.Close()

  $rtf = $word.Documents.Open({rtf_output}, $false, $false)
  $rtf.SaveAs2({migrated_output}, 16)
  $rtf.Close()

  $migrated = $word.Documents.Open({migrated_output}, $false, $true)
  $migrated.SaveAs2({migrated_export}, 2)
  $migratedSignatures = Get-ParagraphSignatureObjects $migrated
  $migrated.Close()
}}
finally {{
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}}

$sourceSignatures | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath {source_signatures}
$migratedSignatures | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath {migrated_signatures}"#,
        staged_input = ps_single_quoted(staged_input),
        source_export = ps_single_quoted(source_export),
        source_signatures = ps_single_quoted(source_signatures),
        rtf_output = ps_single_quoted(rtf_output),
        migrated_output = ps_single_quoted(migrated_output),
        migrated_export = ps_single_quoted(migrated_export),
        migrated_signatures = ps_single_quoted(migrated_signatures),
    );

    fs::write(script_path, script)
        .with_context(|| format!("failed to write migration script {}", script_path.display()))
}

fn ps_single_quoted(path: &Path) -> String {
    format!("'{}'", path.display().to_string().replace('\'', "''"))
}

fn run_powershell_file(script_path: &Path) -> Result<()> {
    let output = Command::new("powershell")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
        ])
        .arg(script_path)
        .output()
        .with_context(|| format!("failed to launch PowerShell for {}", script_path.display()))?;

    if output.status.success() {
        return Ok(());
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    bail!(
        "PowerShell migration script failed (exit code {:?})\nSTDOUT:\n{}\nSTDERR:\n{}",
        output.status.code(),
        stdout.trim(),
        stderr.trim()
    );
}

fn validate_word_fidelity_exports(
    source_signatures_path: &Path,
    migrated_signatures_path: &Path,
) -> Result<FidelityReport> {
    let source_raw = fs::read_to_string(source_signatures_path).with_context(|| {
        format!(
            "failed to read source Word fidelity export {}",
            source_signatures_path.display()
        )
    })?;
    let migrated_raw = fs::read_to_string(migrated_signatures_path).with_context(|| {
        format!(
            "failed to read migrated Word fidelity export {}",
            migrated_signatures_path.display()
        )
    })?;

    let source_signatures: Vec<WordParagraphSignature> =
        parse_word_signature_json(&source_raw).context("failed to parse source Word fidelity export")?;
    let migrated_signatures: Vec<WordParagraphSignature> = parse_word_signature_json(&migrated_raw)
        .context("failed to parse migrated Word fidelity export")?;

    let mut source_counts = std::collections::BTreeMap::<WordParagraphSignature, usize>::new();
    let mut migrated_counts = std::collections::BTreeMap::<WordParagraphSignature, usize>::new();

    for signature in source_signatures {
        *source_counts.entry(signature).or_insert(0usize) += 1;
    }
    for signature in migrated_signatures {
        *migrated_counts.entry(signature).or_insert(0usize) += 1;
    }

    let mut issues = Vec::new();
    for (signature, source_count) in source_counts {
        let migrated_count = migrated_counts.get(&signature).copied().unwrap_or(0usize);
        if migrated_count < source_count {
            issues.push(format!(
                "Word fidelity mismatch for paragraph '{}' (style: {:?}, highlights: {:?}); expected at least {}, found {}",
                signature.text, signature.style, signature.highlights, source_count, migrated_count
            ));
        }
    }

    Ok(FidelityReport { issues })
}

fn parse_word_signature_json(input: &str) -> Result<Vec<WordParagraphSignature>> {
    if input.trim().is_empty() {
        return Ok(Vec::new());
    }

    match serde_json::from_str::<Vec<WordParagraphSignature>>(input) {
        Ok(items) => Ok(items),
        Err(_) => Ok(vec![serde_json::from_str::<WordParagraphSignature>(input)?]),
    }
}

fn format_validation_issues(issues: &[String]) -> String {
    if issues.is_empty() {
        "none".to_string()
    } else {
        issues.join("\n")
    }
}

fn is_story_part(name: &str) -> bool {
    name == DOCUMENT_XML_PATH
        || (name.starts_with("word/header") && name.ends_with(".xml"))
        || (name.starts_with("word/footer") && name.ends_with(".xml"))
}

pub fn apply_spec_to_docx_bytes(
    input_bytes: &[u8],
    spec: &AutomationSpec,
    spec_base_dir: &Path,
) -> Result<Vec<u8>> {
    spec.validate(spec_base_dir)?;
    if can_use_raw_document_insert_path(spec)? {
        return apply_raw_document_insert_path(input_bytes, spec, spec_base_dir);
    }

    let mut package = DocxPackage::read(input_bytes)?;

    for operation in &spec.operations {
        match operation {
            Operation::InsertParagraphs {
                part,
                anchor,
                entries,
            } => {
                let anchor = anchor.as_spec();
                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    for (offset, entry) in entries.iter().enumerate() {
                        story.children.insert(
                            anchor_index + 1 + offset,
                            XMLNode::Element(make_text_paragraph(entry)?),
                        );
                    }
                    Ok(())
                })?;
            }
            Operation::ReplaceText {
                part,
                find,
                replace,
                highlight,
            } => {
                let part_name = part.resolve()?;
                with_xml_part_mut(&mut package, &part_name, |root| {
                    replace_text_in_element(root, find, replace, highlight);
                    Ok(())
                })?;
            }
            Operation::DeleteParagraphs { part, contains } => {
                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    story.children.retain(|node| {
                        !matches!(node, XMLNode::Element(element) if local_name(&element.name) == "p" && paragraph_text(element).contains(contains))
                    });
                    Ok(())
                })?;
            }
            Operation::InsertTableAfter {
                part,
                anchor,
                table,
            } => {
                let anchor = anchor.as_spec();
                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    story
                        .children
                        .insert(anchor_index + 1, XMLNode::Element(make_table(table)?));
                    Ok(())
                })?;
            }
            Operation::InsertHyperlinkAfter {
                part,
                anchor,
                hyperlink,
            } => {
                let anchor = anchor.as_spec();
                let part_name = part.resolve()?;
                let rel_id = add_relationship(
                    &mut package,
                    &part_name,
                    REL_TYPE_HYPERLINK,
                    &hyperlink.url,
                    Some("External"),
                )?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    story.children.insert(
                        anchor_index + 1,
                        XMLNode::Element(make_hyperlink_paragraph(hyperlink, &rel_id)?),
                    );
                    Ok(())
                })?;
            }
            Operation::InsertImageAfter {
                part,
                anchor,
                image,
            } => {
                let anchor = anchor.as_spec();
                let part_name = part.resolve()?;
                let image_path = spec_base_dir.join(&image.path);
                let image_bytes = fs::read(&image_path)
                    .with_context(|| format!("failed to read image {}", image_path.display()))?;
                let extension = image_path
                    .extension()
                    .and_then(|ext| ext.to_str())
                    .map(|ext| ext.to_ascii_lowercase())
                    .with_context(|| {
                        format!("image missing extension: {}", image_path.display())
                    })?;
                ensure_content_type_for_extension(&mut package, &extension)?;
                let media_target = add_media_file(&mut package, &extension, &image_bytes)?;
                let rel_id = add_relationship(
                    &mut package,
                    &part_name,
                    REL_TYPE_IMAGE,
                    &media_target,
                    None,
                )?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    let drawing_id = next_drawing_id(story);
                    story.children.insert(
                        anchor_index + 1,
                        XMLNode::Element(make_image_paragraph(image, &rel_id, drawing_id)?),
                    );
                    Ok(())
                })?;
            }
            Operation::InsertSectionBreakAfter {
                part,
                anchor,
                break_type,
            } => {
                let anchor = anchor.as_spec();
                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    story.children.insert(
                        anchor_index + 1,
                        XMLNode::Element(make_section_break_paragraph(break_type)?),
                    );
                    Ok(())
                })?;
            }
            Operation::InsertCommentAfter {
                part,
                anchor,
                comment,
            } => {
                let anchor = anchor.as_spec();
                ensure_comments_part(&mut package)?;
                let comment_id = next_named_id(&package, "word/comments.xml", "comment")?;
                append_comment_entry(&mut package, comment_id, comment)?;

                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    let Some(XMLNode::Element(paragraph)) = story.children.get_mut(anchor_index) else {
                        bail!("anchor paragraph not found at index {anchor_index}");
                    };
                    attach_comment_to_paragraph(paragraph, comment_id)?;
                    Ok(())
                })?;
            }
            Operation::InsertNoteAfter {
                part,
                anchor,
                kind,
                note,
            } => {
                let anchor = anchor.as_spec();
                let note_id = ensure_note_part_and_next_id(&mut package, kind)?;
                append_note_entry(&mut package, kind, note_id, note)?;

                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    story.children.insert(
                        anchor_index + 1,
                        XMLNode::Element(make_note_reference_paragraph(kind, note, note_id)?),
                    );
                    Ok(())
                })?;
            }
            Operation::InsertContentControlAfter {
                part,
                anchor,
                control,
            } => {
                let anchor = anchor.as_spec();
                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    let control_id = next_sdt_id(story);
                    story.children.insert(
                        anchor_index + 1,
                        XMLNode::Element(make_content_control(control, control_id)?),
                    );
                    Ok(())
                })?;
            }
            Operation::InsertFieldAfter {
                part,
                anchor,
                field,
            } => {
                let anchor = anchor.as_spec();
                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    story.children.insert(
                        anchor_index + 1,
                        XMLNode::Element(make_field_paragraph(field)?),
                    );
                    Ok(())
                })?;
            }
            Operation::TrackInsertParagraphs {
                part,
                anchor,
                author,
                date,
                entries,
            } => {
                let anchor = anchor.as_spec();
                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let anchor_index = find_anchor_index(story, &anchor)?;
                    let mut change_id = next_change_id(story);
                    for (offset, entry) in entries.iter().enumerate() {
                        story.children.insert(
                            anchor_index + 1 + offset,
                            XMLNode::Element(make_tracked_insert_paragraph(
                                entry, change_id, author, date,
                            )?),
                        );
                        change_id += 1;
                    }
                    Ok(())
                })?;
            }
            Operation::UpsertHeaderFooter {
                kind,
                reference,
                section_index,
                entries,
            } => {
                upsert_header_footer(&mut package, kind, reference, *section_index, entries)?;
            }
            Operation::TrackDeleteParagraphs {
                part,
                contains,
                author,
                date,
            } => {
                let part_name = part.resolve()?;
                with_story_part_mut(&mut package, &part_name, |story| {
                    let mut change_id = next_change_id(story);
                    for node in &mut story.children {
                        if let XMLNode::Element(element) = node
                            && local_name(&element.name) == "p"
                            && paragraph_text(element).contains(contains)
                        {
                            let deleted_text = paragraph_text(element);
                            *node = XMLNode::Element(make_tracked_delete_paragraph(
                                &deleted_text,
                                change_id,
                                author,
                                date,
                            )?);
                            change_id += 1;
                        }
                    }
                    Ok(())
                })?;
            }
            Operation::SetCoreProperty { property, value } => {
                ensure_core_properties_part(&mut package)?;
                with_xml_part_mut(&mut package, "docProps/core.xml", |root| {
                    set_core_property(root, property, value);
                    Ok(())
                })?;
            }
        }
    }

    package.write()
}

fn can_use_raw_document_insert_path(spec: &AutomationSpec) -> Result<bool> {
    if spec.operations.is_empty() {
        return Ok(false);
    }

    for operation in &spec.operations {
        match operation {
            Operation::InsertParagraphs { part, .. } | Operation::InsertImageAfter { part, .. } => {
                let resolved = part.resolve()?;
                if resolved != DOCUMENT_XML_PATH {
                    return Ok(false);
                }
            }
            _ => return Ok(false),
        }
    }

    Ok(true)
}

fn apply_raw_document_insert_path(
    input_bytes: &[u8],
    spec: &AutomationSpec,
    spec_base_dir: &Path,
) -> Result<Vec<u8>> {
    let mut package = DocxPackage::read(input_bytes)?;
    let xml_bytes = package
        .get_file(DOCUMENT_XML_PATH)
        .context("missing document.xml")?
        .to_vec();
    let xml = String::from_utf8(xml_bytes).context("document.xml is not utf-8")?;
    let updated = apply_raw_insert_operations_to_document_xml(&mut package, &xml, spec, spec_base_dir)?;
    package.set_file(DOCUMENT_XML_PATH, updated.into_bytes());
    package.write()
}

fn apply_raw_insert_operations_to_document_xml(
    package: &mut DocxPackage,
    xml: &str,
    spec: &AutomationSpec,
    spec_base_dir: &Path,
) -> Result<String> {
    let paragraph_re = Regex::new(r"(?s)<w:p\b.*?</w:p>").context("invalid paragraph regex")?;
    let tag_re = Regex::new(r"(?s)<[^>]+>").context("invalid tag regex")?;
    let root = parse_xml(xml.as_bytes())?;
    let story = story_container_ref(&root)?;
    let mut next_image_drawing_id = next_drawing_id(story);

    #[derive(Debug)]
    struct ParagraphMatch {
        end: usize,
        text: String,
    }

    let paragraphs: Vec<ParagraphMatch> = paragraph_re
        .find_iter(xml)
        .map(|m| ParagraphMatch {
            end: m.end(),
            text: decode_xml_entities(&tag_re.replace_all(m.as_str(), "").into_owned()),
        })
        .collect();

    #[derive(Debug)]
    struct Insertion {
        pos: usize,
        order: usize,
        xml: String,
    }

    let mut insertions = Vec::new();
    for (order, operation) in spec.operations.iter().enumerate() {
        let anchor = match operation {
            Operation::InsertParagraphs { anchor, .. } | Operation::InsertImageAfter { anchor, .. } => {
                anchor.as_spec()
            }
            _ => bail!("raw document insert path only supports insert-paragraphs and insert-image-after"),
        };
        let mut matches = 0usize;
        let mut insertion_pos = None;
        for paragraph in &paragraphs {
            if anchor_matches(&paragraph.text, &anchor) {
                matches += 1;
                if matches == anchor.occurrence {
                    insertion_pos = Some(paragraph.end);
                    break;
                }
            }
        }
        let Some(pos) = insertion_pos else {
            bail!(
                "anchor not found: '{}' occurrence {}",
                anchor.text,
                anchor.occurrence
            );
        };
        let fragment = match operation {
            Operation::InsertParagraphs { entries, .. } => {
                entries.iter().map(raw_text_paragraph_xml).collect::<String>()
            }
            Operation::InsertImageAfter { image, .. } => {
                let image_path = spec_base_dir.join(&image.path);
                let image_bytes = fs::read(&image_path)
                    .with_context(|| format!("failed to read image {}", image_path.display()))?;
                let extension = image_path
                    .extension()
                    .and_then(|ext| ext.to_str())
                    .map(|ext| ext.to_ascii_lowercase())
                    .with_context(|| format!("image missing extension: {}", image_path.display()))?;
                ensure_content_type_for_extension(package, &extension)?;
                let media_target = add_media_file(package, &extension, &image_bytes)?;
                let rel_id =
                    add_relationship(package, DOCUMENT_XML_PATH, REL_TYPE_IMAGE, &media_target, None)?;
                let drawing_id = next_image_drawing_id;
                next_image_drawing_id += 1;
                raw_image_paragraph_xml(image, &rel_id, drawing_id)
            }
            _ => unreachable!(),
        };
        insertions.push(Insertion {
            pos,
            order,
            xml: fragment,
        });
    }

    insertions.sort_by(|a, b| b.pos.cmp(&a.pos).then(b.order.cmp(&a.order)));

    let mut output = xml.to_string();
    for insertion in insertions {
        output.insert_str(insertion.pos, &insertion.xml);
    }

    Ok(output)
}

fn resolve_part_name(value: Option<&str>) -> Result<String> {
    match value.unwrap_or("document") {
        "document" => Ok(DOCUMENT_XML_PATH.to_string()),
        explicit if explicit.starts_with("word/") => Ok(explicit.to_string()),
        explicit if explicit.starts_with("header:") => {
            let suffix = explicit.trim_start_matches("header:");
            Ok(format!("word/header{suffix}.xml"))
        }
        explicit if explicit.starts_with("footer:") => {
            let suffix = explicit.trim_start_matches("footer:");
            Ok(format!("word/footer{suffix}.xml"))
        }
        unsupported => bail!("unsupported part target: {unsupported}"),
    }
}

fn with_story_part_mut(
    package: &mut DocxPackage,
    part_name: &str,
    operation: impl FnOnce(&mut Element) -> Result<()>,
) -> Result<()> {
    with_xml_part_mut(package, part_name, |root| {
        let story = story_container_mut(root)?;
        operation(story)
    })
}

fn with_xml_part_mut(
    package: &mut DocxPackage,
    part_name: &str,
    operation: impl FnOnce(&mut Element) -> Result<()>,
) -> Result<()> {
    let bytes = package
        .get_file(part_name)
        .with_context(|| format!("missing part {part_name}"))?
        .to_vec();
    let mut root = parse_xml(&bytes)?;
    operation(&mut root)?;
    ensure_root_namespaces(&mut root);
    let rewritten = write_xml(&root)?;
    package.set_file(part_name, rewritten);
    Ok(())
}

fn ensure_root_namespaces(root: &mut Element) {
    if local_name(&root.name) != "document" {
        return;
    }

    let namespaces = root.namespaces.get_or_insert_with(Namespace::empty);
    if namespaces.get("r").is_none() {
        namespaces.put("r".to_string(), R_NS.to_string());
    }
}

fn story_container_mut(root: &mut Element) -> Result<&mut Element> {
    match local_name(&root.name) {
        "document" => find_child_mut_local(root, "body").context("document missing body"),
        "hdr" | "ftr" => Ok(root),
        other => bail!("unsupported story root {other}"),
    }
}

fn story_container_ref(root: &Element) -> Result<&Element> {
    match local_name(&root.name) {
        "document" => find_child_ref_local(root, "body").context("document missing body"),
        "hdr" | "ftr" => Ok(root),
        other => bail!("unsupported story root {other}"),
    }
}

fn find_anchor_index(story: &Element, anchor: &AnchorSpec) -> Result<usize> {
    let mut matched_count = 0usize;
    story
        .children
        .iter()
        .enumerate()
        .find_map(|(idx, node)| match node {
            XMLNode::Element(element)
                if local_name(&element.name) == "p"
                    && anchor_matches(&paragraph_text(element), anchor) =>
            {
                matched_count += 1;
                (matched_count == anchor.occurrence).then_some(idx)
            }
            _ => None,
        })
        .with_context(|| {
            format!(
                "anchor not found: '{}' occurrence {}",
                anchor.text, anchor.occurrence
            )
        })
}

fn anchor_matches(text: &str, anchor: &AnchorSpec) -> bool {
    match anchor.mode {
        AnchorMatchMode::Contains => text.contains(&anchor.text),
        AnchorMatchMode::Equals => text == anchor.text,
        AnchorMatchMode::StartsWith => text.starts_with(&anchor.text),
        AnchorMatchMode::EndsWith => text.ends_with(&anchor.text),
    }
}

fn paragraph_text(paragraph: &Element) -> String {
    let mut output = String::new();
    collect_text(paragraph, &mut output);
    output
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
struct ParagraphSignature {
    text: String,
    style: Option<String>,
    highlights: Vec<String>,
}

fn paragraph_signatures(
    story: &Element,
    skip: impl Fn(&Element) -> bool,
) -> std::collections::BTreeMap<ParagraphSignature, usize> {
    let mut counts = std::collections::BTreeMap::new();
    for child in &story.children {
        let XMLNode::Element(paragraph) = child else {
            continue;
        };
        if local_name(&paragraph.name) != "p" || skip(paragraph) {
            continue;
        }
        let signature = ParagraphSignature {
            text: paragraph_text(paragraph),
            style: paragraph_style(paragraph),
            highlights: paragraph_highlights(paragraph),
        };
        *counts.entry(signature).or_insert(0usize) += 1;
    }
    counts
}

fn paragraph_style(paragraph: &Element) -> Option<String> {
    find_child_ref_local(paragraph, "pPr")
        .and_then(|ppr| find_child_ref_local(ppr, "pStyle"))
        .and_then(|style| {
            style
                .attributes
                .get("w:val")
                .cloned()
                .or_else(|| style.attributes.get("val").cloned())
        })
}

fn paragraph_highlights(paragraph: &Element) -> Vec<String> {
    let mut values = Vec::new();
    collect_highlights(paragraph, &mut values);
    values
}

fn collect_highlights(element: &Element, output: &mut Vec<String>) {
    if local_name(&element.name) == "highlight" {
        if let Some(value) = element
            .attributes
            .get("w:val")
            .or_else(|| element.attributes.get("val"))
        {
            output.push(value.clone());
        }
    }

    for child in &element.children {
        if let XMLNode::Element(element) = child {
            collect_highlights(element, output);
        }
    }
}

pub fn prepare_work_session(
    input: &Path,
    session_id: &str,
    temp_root: Option<&Path>,
    trusted_ooxml: Option<&Path>,
    allow_word_com_encrypted_package: bool,
) -> Result<WorkSessionReport> {
    ensure_session_id_is_safe(session_id)?;
    let temp_root = temp_root.unwrap_or(Path::new(DEFAULT_TEMP_ROOT));
    let session_dir = session_dir(temp_root, session_id);
    fs::create_dir_all(&session_dir)
        .with_context(|| format!("failed to create session dir {}", session_dir.display()))?;

    let input_bytes = fs::read(input).with_context(|| format!("failed to read {}", input.display()))?;
    let source_sha256 = sha256_hex(&input_bytes);
    let normalized_path = session_dir.join(SESSION_WORKING_DOCX_FILE);

    if let Ok(metadata) = load_work_session_metadata(temp_root, session_id)
        && metadata.source_path == input.to_string_lossy()
        && metadata.source_sha256 == source_sha256
        && Path::new(&metadata.normalized_path).exists()
    {
        return Ok(WorkSessionReport {
            session_id: session_id.to_string(),
            session_dir,
            normalized_input: PathBuf::from(metadata.normalized_path),
            detected_format: metadata.detected_format,
            cache_hit: true,
        });
    }

    let normalization = normalize_docx_file(
        input,
        &normalized_path,
        Some(temp_root),
        trusted_ooxml,
        allow_word_com_encrypted_package,
    )?;
    let output_dir = input
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));
    let metadata = WorkSessionMetadata {
        session_id: session_id.to_string(),
        source_path: input.to_string_lossy().to_string(),
        source_sha256,
        normalized_path: normalized_path.to_string_lossy().to_string(),
        output_dir: output_dir.to_string_lossy().to_string(),
        current_version_path: input.to_string_lossy().to_string(),
        detected_format: normalization.detected_format,
    };
    save_work_session_metadata(temp_root, &metadata)?;

    Ok(WorkSessionReport {
        session_id: session_id.to_string(),
        session_dir,
        normalized_input: normalized_path,
        detected_format: normalization.detected_format,
        cache_hit: false,
    })
}

pub fn publish_spec_file_to_next_version(
    input: &Path,
    spec: &AutomationSpec,
    spec_base_dir: &Path,
    temp_root: Option<&Path>,
    output_dir: Option<&Path>,
    mode: PublishTargetMode,
) -> Result<PublishNextWorkflowReport> {
    let (resolved_output, version_number) = resolve_versioned_output_path(input, output_dir, mode)?;
    let report = publish_spec_file_to_docx_internal(
        input,
        &resolved_output,
        spec,
        spec_base_dir,
        temp_root,
        mode == PublishTargetMode::Latest,
    )?;
    Ok(PublishNextWorkflowReport {
        xml_parts_checked: report.xml_parts_checked,
        published_output: report.published_output,
        version_number,
        mode,
    })
}

pub fn publish_session_to_next_version(
    session_id: &str,
    spec: &AutomationSpec,
    spec_base_dir: &Path,
    temp_root: Option<&Path>,
    output_dir: Option<&Path>,
    mode: PublishTargetMode,
) -> Result<PublishNextWorkflowReport> {
    ensure_session_id_is_safe(session_id)?;
    let temp_root = temp_root.unwrap_or(Path::new(DEFAULT_TEMP_ROOT));
    let mut metadata = load_work_session_metadata(temp_root, session_id)?;
    let working_input = PathBuf::from(&metadata.normalized_path);
    let source_bytes = fs::read(&working_input)
        .with_context(|| format!("failed to read session working copy {}", working_input.display()))?;
    let (output_bytes, candidate_validation) =
        build_validated_candidate_from_source_bytes(&working_input, &source_bytes, spec, spec_base_dir)?;

    let current_version_path = PathBuf::from(&metadata.current_version_path);
    let requested_output_dir = output_dir.map(Path::to_path_buf);
    let fallback_output_dir = PathBuf::from(&metadata.output_dir);
    let (resolved_output, version_number) = resolve_versioned_output_path(
        &current_version_path,
        requested_output_dir.as_deref().or(Some(&fallback_output_dir)),
        mode,
    )?;

    if let Some(parent) = resolved_output.parent()
        && !parent.as_os_str().is_empty()
    {
        fs::create_dir_all(parent)
            .with_context(|| format!("failed to create output directory {}", parent.display()))?;
    }
    ensure_existing_output_is_safe_for_direct_write(&resolved_output)?;
    fs::write(&resolved_output, &output_bytes)
        .with_context(|| format!("failed to publish {}", resolved_output.display()))?;
    verify_published_output_matches_candidate(&resolved_output, &output_bytes)?;
    fs::write(&working_input, &output_bytes).with_context(|| {
        format!(
            "failed to refresh session working copy {}",
            working_input.display()
        )
    })?;

    metadata.current_version_path = resolved_output.to_string_lossy().to_string();
    metadata.source_sha256 = sha256_hex(&output_bytes);
    if let Some(parent) = resolved_output.parent() {
        metadata.output_dir = parent.to_string_lossy().to_string();
    }
    save_work_session_metadata(temp_root, &metadata)?;

    Ok(PublishNextWorkflowReport {
        xml_parts_checked: candidate_validation.xml_parts_checked,
        published_output: resolved_output,
        version_number,
        mode,
    })
}

fn should_skip_paragraph_fidelity_check(part_name: &str, paragraph: &Element, spec: &AutomationSpec) -> bool {
    let text = paragraph_text(paragraph);
    spec.operations.iter().any(|operation| operation_touches_existing_paragraph(operation, part_name, &text))
}

fn operation_touches_existing_paragraph(operation: &Operation, part_name: &str, paragraph_text: &str) -> bool {
    match operation {
        Operation::ReplaceText { part, find, .. } => {
            part_matches_target(part, part_name) && paragraph_text.contains(find)
        }
        Operation::DeleteParagraphs { part, contains }
        | Operation::TrackDeleteParagraphs { part, contains, .. } => {
            part_matches_target(part, part_name) && paragraph_text.contains(contains)
        }
        _ => false,
    }
}

fn part_matches_target(part: &PartTarget, part_name: &str) -> bool {
    part.resolve().map(|resolved| resolved == part_name).unwrap_or(false)
}

fn collect_text(element: &Element, output: &mut String) {
    if local_name(&element.name) == "t" {
        for child in &element.children {
            if let XMLNode::Text(text) = child {
                output.push_str(text);
            }
        }
    }

    for child in &element.children {
        if let XMLNode::Element(element) = child {
            collect_text(element, output);
        }
    }
}

fn replace_text_in_element(
    element: &mut Element,
    find: &str,
    replace: &str,
    highlight: &str,
) -> bool {
    let mut changed = false;

    if local_name(&element.name) == "r" {
        let mut run_changed = false;
        for child in &mut element.children {
            if let XMLNode::Element(text_element) = child
                && local_name(&text_element.name) == "t"
            {
                for text_node in &mut text_element.children {
                    if let XMLNode::Text(text) = text_node
                        && text.contains(find)
                    {
                        *text = text.replace(find, replace);
                        run_changed = true;
                        changed = true;
                    }
                }
            }
        }

        if run_changed {
            ensure_run_highlight(element, highlight);
        }
    }

    for child in &mut element.children {
        if let XMLNode::Element(element) = child
            && replace_text_in_element(element, find, replace, highlight)
        {
            changed = true;
        }
    }

    changed
}

fn ensure_run_highlight(run: &mut Element, highlight: &str) {
    let rpr_index = run.children.iter().position(
        |child| matches!(child, XMLNode::Element(element) if local_name(&element.name) == "rPr"),
    );

    let index = if let Some(index) = rpr_index {
        index
    } else {
        run.children
            .insert(0, XMLNode::Element(Element::new("w:rPr")));
        0
    };

    let XMLNode::Element(rpr) = &mut run.children[index] else {
        return;
    };

    if !rpr
        .children
        .iter()
        .any(|child| matches!(child, XMLNode::Element(element) if local_name(&element.name) == "highlight"))
    {
        let mut highlight_element = Element::new("w:highlight");
        highlight_element
            .attributes
            .insert("w:val".to_string(), highlight.to_string());
        rpr.children.push(XMLNode::Element(highlight_element));
    }
}

fn next_drawing_id(story: &Element) -> u32 {
    let mut max_id = 0u32;
    collect_max_drawing_id(story, &mut max_id);
    max_id + 1
}

fn collect_max_drawing_id(element: &Element, max_id: &mut u32) {
    if local_name(&element.name) == "docPr"
        && let Some(id_value) = element.attributes.get("id")
        && let Ok(parsed) = id_value.parse::<u32>()
    {
        *max_id = (*max_id).max(parsed);
    }

    for child in &element.children {
        if let XMLNode::Element(element) = child {
            collect_max_drawing_id(element, max_id);
        }
    }
}

fn set_core_property(root: &mut Element, property: &CoreProperty, value: &str) {
    if let Some(existing) = find_child_mut_local(root, property.local_name()) {
        existing.children.clear();
        existing.children.push(XMLNode::Text(value.to_string()));
        return;
    }

    let mut property_element = Element::new(property.xml_name());
    property_element
        .children
        .push(XMLNode::Text(value.to_string()));
    root.children.push(XMLNode::Element(property_element));
}

fn ensure_core_properties_part(package: &mut DocxPackage) -> Result<()> {
    if package.get_file("docProps/core.xml").is_none() {
        package.ensure_directory("docProps/");
        package.set_file("docProps/core.xml", minimal_core_properties_xml());
        ensure_root_relationship(package, REL_TYPE_CORE_PROPS, "docProps/core.xml")?;
        ensure_core_content_type(package)?;
    }
    Ok(())
}

fn ensure_root_relationship(package: &mut DocxPackage, rel_type: &str, target: &str) -> Result<()> {
    let bytes = package
        .get_file(ROOT_RELS_XML_PATH)
        .map(|bytes| bytes.to_vec())
        .unwrap_or_else(minimal_root_relationships_xml);
    let mut root = parse_xml(&bytes)?;

    if !root.children.iter().any(|child| {
        matches!(child, XMLNode::Element(element)
            if element.name == "Relationship"
            && element.attributes.get("Type").map(|value| value.as_str()) == Some(rel_type)
            && element.attributes.get("Target").map(|value| value.as_str()) == Some(target))
    }) {
        let next_id = next_relationship_id(&root);
        let mut relationship = Element::new("Relationship");
        relationship
            .attributes
            .insert("Id".to_string(), format!("rId{next_id}"));
        relationship
            .attributes
            .insert("Type".to_string(), rel_type.to_string());
        relationship
            .attributes
            .insert("Target".to_string(), target.to_string());
        root.children.push(XMLNode::Element(relationship));
    }

    package.ensure_directory("_rels/");
    package.set_file(ROOT_RELS_XML_PATH, write_xml(&root)?);
    Ok(())
}

fn add_relationship(
    package: &mut DocxPackage,
    part_name: &str,
    rel_type: &str,
    target: &str,
    target_mode: Option<&str>,
) -> Result<String> {
    let relationships_path = part_relationships_path(part_name)?;
    let xml_bytes = package
        .get_file(&relationships_path)
        .map(|bytes| bytes.to_vec())
        .unwrap_or_else(minimal_part_relationships_xml);
    let mut root = parse_xml(&xml_bytes)?;
    let next_id = format!("rId{}", next_relationship_id(&root));
    let mut relationship = Element::new("Relationship");
    relationship
        .attributes
        .insert("Id".to_string(), next_id.clone());
    relationship
        .attributes
        .insert("Type".to_string(), rel_type.to_string());
    relationship
        .attributes
        .insert("Target".to_string(), target.to_string());
    if let Some(mode) = target_mode {
        relationship
            .attributes
            .insert("TargetMode".to_string(), mode.to_string());
    }
    root.children.push(XMLNode::Element(relationship));

    if let Some(parent) = Path::new(&relationships_path).parent() {
        let folder = format!("{}\\", parent.display()).replace('\\', "/");
        package.ensure_directory(&folder);
    }
    package.set_file(&relationships_path, write_xml(&root)?);
    Ok(next_id)
}

fn part_relationships_path(part_name: &str) -> Result<String> {
    let path = Path::new(part_name);
    let file_name = path
        .file_name()
        .and_then(|value| value.to_str())
        .with_context(|| format!("invalid part name {part_name}"))?;
    let parent = path
        .parent()
        .and_then(|value| value.to_str())
        .with_context(|| format!("invalid part parent {part_name}"))?;
    Ok(format!("{parent}/_rels/{file_name}.rels"))
}

fn next_relationship_id(root: &Element) -> u32 {
    let mut max_id = 0u32;
    for child in &root.children {
        if let XMLNode::Element(element) = child
            && element.name == "Relationship"
            && let Some(id) = element.attributes.get("Id")
            && let Some(suffix) = id.strip_prefix("rId")
            && let Ok(parsed) = suffix.parse::<u32>()
        {
            max_id = max_id.max(parsed);
        }
    }
    max_id + 1
}

fn add_media_file(package: &mut DocxPackage, extension: &str, bytes: &[u8]) -> Result<String> {
    package.ensure_directory("word/media/");

    let mut index = 1usize;
    loop {
        let name = format!("word/media/image{index}.{extension}");
        if package.get_file(&name).is_none() {
            package.set_file(&name, bytes.to_vec());
            return Ok(format!("media/image{index}.{extension}"));
        }
        index += 1;
    }
}

fn ensure_core_content_type(package: &mut DocxPackage) -> Result<()> {
    let bytes = package
        .get_file(CONTENT_TYPES_XML_PATH)
        .context("missing [Content_Types].xml")?
        .to_vec();
    let mut root = parse_xml(&bytes)?;

    let exists = root.children.iter().any(|child| {
        matches!(child, XMLNode::Element(element)
            if element.name == "Override"
            && element.attributes.get("PartName").map(|value| value.as_str()) == Some("/docProps/core.xml"))
    });
    if !exists {
        let mut override_element = Element::new("Override");
        override_element
            .attributes
            .insert("PartName".to_string(), "/docProps/core.xml".to_string());
        override_element.attributes.insert(
            "ContentType".to_string(),
            CORE_PROPS_CONTENT_TYPE.to_string(),
        );
        root.children.push(XMLNode::Element(override_element));
    }

    package.set_file(CONTENT_TYPES_XML_PATH, write_xml(&root)?);
    Ok(())
}

fn ensure_content_type_for_extension(package: &mut DocxPackage, extension: &str) -> Result<()> {
    let content_type = match extension {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        _ => bail!("unsupported image extension: {extension}"),
    };

    let bytes = package
        .get_file(CONTENT_TYPES_XML_PATH)
        .context("missing [Content_Types].xml")?
        .to_vec();
    let mut root = parse_xml(&bytes)?;

    let exists = root.children.iter().any(|child| {
        matches!(child, XMLNode::Element(element)
            if element.name == "Default"
            && element.attributes.get("Extension").map(|value| value.as_str()) == Some(extension))
    });

    if !exists {
        let mut default_element = Element::new("Default");
        default_element
            .attributes
            .insert("Extension".to_string(), extension.to_string());
        default_element
            .attributes
            .insert("ContentType".to_string(), content_type.to_string());
        root.children.push(XMLNode::Element(default_element));
    }

    package.set_file(CONTENT_TYPES_XML_PATH, write_xml(&root)?);
    Ok(())
}

fn ensure_override_content_type(
    package: &mut DocxPackage,
    part_name: &str,
    content_type: &str,
) -> Result<()> {
    let bytes = package
        .get_file(CONTENT_TYPES_XML_PATH)
        .context("missing [Content_Types].xml")?
        .to_vec();
    let mut root = parse_xml(&bytes)?;
    let part_override = format!("/{}", part_name.replace('\\', "/"));

    let exists = root.children.iter().any(|child| {
        matches!(child, XMLNode::Element(element)
            if local_name(&element.name) == "Override"
            && element.attributes.get("PartName").map(|value| value.as_str()) == Some(part_override.as_str()))
    });

    if !exists {
        let mut override_element = Element::new("Override");
        override_element
            .attributes
            .insert("PartName".to_string(), part_override);
        override_element
            .attributes
            .insert("ContentType".to_string(), content_type.to_string());
        root.children.push(XMLNode::Element(override_element));
    }

    package.set_file(CONTENT_TYPES_XML_PATH, write_xml(&root)?);
    Ok(())
}

fn ensure_comments_part(package: &mut DocxPackage) -> Result<()> {
    if package.get_file("word/comments.xml").is_none() {
        package.set_file("word/comments.xml", minimal_comments_xml());
        ensure_override_content_type(package, "word/comments.xml", COMMENTS_CONTENT_TYPE)?;
        add_relationship(
            package,
            DOCUMENT_XML_PATH,
            REL_TYPE_COMMENTS,
            "comments.xml",
            None,
        )?;
    }
    if package.get_file("word/commentsExtended.xml").is_none() {
        package.set_file("word/commentsExtended.xml", minimal_comments_extended_xml());
        ensure_override_content_type(
            package,
            "word/commentsExtended.xml",
            COMMENTS_EXTENDED_CONTENT_TYPE,
        )?;
        add_relationship(
            package,
            DOCUMENT_XML_PATH,
            REL_TYPE_COMMENTS_EXTENDED,
            "commentsExtended.xml",
            None,
        )?;
    }
    if package.get_file("word/commentsIds.xml").is_none() {
        package.set_file("word/commentsIds.xml", minimal_comments_ids_xml());
        ensure_override_content_type(package, "word/commentsIds.xml", COMMENTS_IDS_CONTENT_TYPE)?;
        add_relationship(
            package,
            DOCUMENT_XML_PATH,
            REL_TYPE_COMMENTS_IDS,
            "commentsIds.xml",
            None,
        )?;
    }
    if package.get_file("word/people.xml").is_none() {
        package.set_file("word/people.xml", minimal_people_xml());
        ensure_override_content_type(package, "word/people.xml", PEOPLE_CONTENT_TYPE)?;
        add_relationship(
            package,
            DOCUMENT_XML_PATH,
            REL_TYPE_PEOPLE,
            "people.xml",
            None,
        )?;
    }
    Ok(())
}

fn upsert_header_footer(
    package: &mut DocxPackage,
    kind: &HeaderFooterKind,
    reference: &HeaderFooterReferenceKind,
    section_index: usize,
    entries: &[ParagraphEntry],
) -> Result<()> {
    let (part_prefix, rel_type, content_type, root_name) = match kind {
        HeaderFooterKind::Header => ("header", REL_TYPE_HEADER, HEADER_CONTENT_TYPE, "w:hdr"),
        HeaderFooterKind::Footer => ("footer", REL_TYPE_FOOTER, FOOTER_CONTENT_TYPE, "w:ftr"),
    };

    let part_name = next_or_existing_header_footer_part(package, kind, section_index)?;
    if package.get_file(&part_name).is_none() {
        package.set_file(&part_name, minimal_header_footer_xml(root_name));
        ensure_override_content_type(package, &part_name, content_type)?;
    }

    with_xml_part_mut(package, &part_name, |root| {
        root.children.clear();
        for entry in entries {
            root.children
                .push(XMLNode::Element(make_text_paragraph(entry)?));
        }
        if root.children.is_empty() {
            root.children
                .push(XMLNode::Element(parse_wrapped_fragment("<w:p />")?));
        }
        Ok(())
    })?;

    let target_name = format!("{part_prefix}{}.xml", part_name_number(&part_name)?);
    let rel_id = add_relationship(package, DOCUMENT_XML_PATH, rel_type, &target_name, None)?;
    attach_header_footer_reference(package, kind, reference, section_index, &rel_id)
}

fn next_or_existing_header_footer_part(
    package: &DocxPackage,
    kind: &HeaderFooterKind,
    section_index: usize,
) -> Result<String> {
    let document_bytes = package
        .get_file(DOCUMENT_XML_PATH)
        .context("missing document.xml")?;
    let root = parse_xml(document_bytes)?;
    let body = story_container_ref(&root)?;
    if let Some(section) = section_at_index(body, section_index)
        && let Some(existing_rel_id) = find_header_footer_rel_id(section, kind)
        && let Some(target) =
            lookup_relationship_target(package, DOCUMENT_XML_PATH, &existing_rel_id)?
    {
        let normalized = if target.starts_with("word/") {
            target
        } else {
            format!("word/{}", target)
        };
        return Ok(normalized);
    }

    let prefix = match kind {
        HeaderFooterKind::Header => "word/header",
        HeaderFooterKind::Footer => "word/footer",
    };
    let mut next_num = 1usize;
    for entry in &package.entries {
        if entry.name.starts_with(prefix)
            && entry.name.ends_with(".xml")
            && let Some(stem) = entry
                .name
                .trim_start_matches(prefix)
                .trim_end_matches(".xml")
                .parse::<usize>()
                .ok()
        {
            next_num = next_num.max(stem + 1);
        }
    }
    Ok(format!("{prefix}{next_num}.xml"))
}

fn attach_header_footer_reference(
    package: &mut DocxPackage,
    kind: &HeaderFooterKind,
    reference: &HeaderFooterReferenceKind,
    section_index: usize,
    rel_id: &str,
) -> Result<()> {
    with_xml_part_mut(package, DOCUMENT_XML_PATH, |root| {
        let body = story_container_mut(root)?;
        ensure_section_count(body, section_index + 1)?;
        let section = section_at_index_mut(body, section_index).context("missing section")?;
        let sect_pr = ensure_sect_pr_mut(section)?;

        let ref_local = match kind {
            HeaderFooterKind::Header => "headerReference",
            HeaderFooterKind::Footer => "footerReference",
        };
        let ref_name = match kind {
            HeaderFooterKind::Header => "w:headerReference",
            HeaderFooterKind::Footer => "w:footerReference",
        };
        let ref_type = match reference {
            HeaderFooterReferenceKind::Default => "default",
            HeaderFooterReferenceKind::First => "first",
            HeaderFooterReferenceKind::Even => "even",
        };

        sect_pr.children.retain(|child| {
            !matches!(child, XMLNode::Element(element)
                if local_name(&element.name) == ref_local
                && element.attributes.get("w:type").or_else(|| element.attributes.get("type")).map(|v| v.as_str()) == Some(ref_type))
        });

        let mut reference_element = Element::new(ref_name);
        reference_element
            .attributes
            .insert("w:type".to_string(), ref_type.to_string());
        reference_element
            .attributes
            .insert("r:id".to_string(), rel_id.to_string());
        sect_pr.children.push(XMLNode::Element(reference_element));
        Ok(())
    })
}

fn ensure_section_count(body: &mut Element, count: usize) -> Result<()> {
    while sections_in_body(body) < count {
        body.children.push(XMLNode::Element(parse_wrapped_fragment(
            "<w:p><w:pPr><w:sectPr /></w:pPr></w:p>",
        )?));
    }
    Ok(())
}

fn sections_in_body(body: &Element) -> usize {
    let mut count = 0usize;
    for child in &body.children {
        if let XMLNode::Element(element) = child
            && (has_local_child(element, "sectPr") || (local_name(&element.name) == "sectPr"))
        {
            count += 1;
        }
    }
    count.max(1)
}

fn section_at_index<'a>(body: &'a Element, index: usize) -> Option<&'a Element> {
    let mut current = 0usize;
    for child in &body.children {
        if let XMLNode::Element(element) = child {
            if has_local_child(element, "sectPr") || local_name(&element.name) == "sectPr" {
                if current == index {
                    return Some(element);
                }
                current += 1;
            }
        }
    }
    None
}

fn section_at_index_mut<'a>(body: &'a mut Element, index: usize) -> Option<&'a mut Element> {
    let mut current = 0usize;
    for child in &mut body.children {
        if let XMLNode::Element(element) = child {
            if has_local_child(element, "sectPr") || local_name(&element.name) == "sectPr" {
                if current == index {
                    return Some(element);
                }
                current += 1;
            }
        }
    }
    None
}

fn ensure_sect_pr_mut(section: &mut Element) -> Result<&mut Element> {
    if local_name(&section.name) == "sectPr" {
        return Ok(section);
    }
    let ppr_index = if let Some(index) = section.children.iter().position(
        |child| matches!(child, XMLNode::Element(element) if local_name(&element.name) == "pPr"),
    ) {
        index
    } else {
        section
            .children
            .insert(0, XMLNode::Element(Element::new("w:pPr")));
        0
    };
    let XMLNode::Element(ppr) = &mut section.children[ppr_index] else {
        bail!("missing pPr")
    };
    let sect_index = if let Some(index) = ppr.children.iter().position(
        |child| matches!(child, XMLNode::Element(element) if local_name(&element.name) == "sectPr"),
    ) {
        index
    } else {
        ppr.children
            .push(XMLNode::Element(Element::new("w:sectPr")));
        ppr.children.len() - 1
    };
    let XMLNode::Element(sect_pr) = &mut ppr.children[sect_index] else {
        bail!("missing sectPr")
    };
    Ok(sect_pr)
}

fn find_header_footer_rel_id(section: &Element, kind: &HeaderFooterKind) -> Option<String> {
    let target_local = match kind {
        HeaderFooterKind::Header => "headerReference",
        HeaderFooterKind::Footer => "footerReference",
    };
    let sect_pr = if local_name(&section.name) == "sectPr" {
        Some(section)
    } else {
        find_child_ref_local(section, "pPr").and_then(|ppr| find_child_ref_local(ppr, "sectPr"))
    }?;
    for child in &sect_pr.children {
        if let XMLNode::Element(element) = child
            && local_name(&element.name) == target_local
            && let Some(id) = element
                .attributes
                .get("r:id")
                .or_else(|| element.attributes.get("id"))
        {
            return Some(id.clone());
        }
    }
    None
}

fn lookup_relationship_target(
    package: &DocxPackage,
    part_name: &str,
    rel_id: &str,
) -> Result<Option<String>> {
    let rels_path = part_relationships_path(part_name)?;
    let Some(bytes) = package.get_file(&rels_path) else {
        return Ok(None);
    };
    let root = parse_xml(bytes)?;
    for child in &root.children {
        if let XMLNode::Element(element) = child
            && local_name(&element.name) == "Relationship"
            && element.attributes.get("Id").map(|v| v.as_str()) == Some(rel_id)
        {
            return Ok(element.attributes.get("Target").cloned());
        }
    }
    Ok(None)
}

fn part_name_number(part_name: &str) -> Result<usize> {
    let stem = Path::new(part_name)
        .file_stem()
        .and_then(|s| s.to_str())
        .context("invalid part name")?;
    let digits = stem
        .chars()
        .rev()
        .take_while(|ch| ch.is_ascii_digit())
        .collect::<String>();
    let digits = digits.chars().rev().collect::<String>();
    digits
        .parse::<usize>()
        .context("missing numeric part suffix")
}

fn has_local_child(element: &Element, local: &str) -> bool {
    element
        .children
        .iter()
        .any(|child| matches!(child, XMLNode::Element(inner) if local_name(&inner.name) == local))
}

fn ensure_note_part_and_next_id(package: &mut DocxPackage, kind: &NoteKind) -> Result<u32> {
    let (part_name, rel_type, content_type, initial_xml, target) = match kind {
        NoteKind::Footnote => (
            "word/footnotes.xml",
            REL_TYPE_FOOTNOTES,
            FOOTNOTES_CONTENT_TYPE,
            minimal_footnotes_xml(),
            "footnotes.xml",
        ),
        NoteKind::Endnote => (
            "word/endnotes.xml",
            REL_TYPE_ENDNOTES,
            ENDNOTES_CONTENT_TYPE,
            minimal_endnotes_xml(),
            "endnotes.xml",
        ),
    };

    if package.get_file(part_name).is_none() {
        package.set_file(part_name, initial_xml);
        ensure_override_content_type(package, part_name, content_type)?;
        add_relationship(package, DOCUMENT_XML_PATH, rel_type, target, None)?;
    }

    let element_name = match kind {
        NoteKind::Footnote => "footnote",
        NoteKind::Endnote => "endnote",
    };
    next_named_id(package, part_name, element_name)
}

fn append_comment_entry(
    package: &mut DocxPackage,
    comment_id: u32,
    comment: &CommentSpec,
) -> Result<()> {
    let comment_para_id = next_hex_attr_id(package, "word/comments.xml", "comment", "paraId")?;
    with_xml_part_mut(package, "word/comments.xml", |root| {
        let namespaces = root.namespaces.get_or_insert_with(Namespace::empty);
        if namespaces.get("w15").is_none() {
            namespaces.put("w15".to_string(), W15_NS.to_string());
        }
        root.children
            .push(XMLNode::Element(make_comment_entry(comment, comment_id, &comment_para_id)?));
        Ok(())
    })?;
    append_comment_compatibility_metadata(package, &comment_para_id, &comment.author)
}

fn read_comments_from_package(package: &DocxPackage) -> Result<Vec<CommentRecord>> {
    let Some(comment_bytes) = package.get_file("word/comments.xml") else {
        return Ok(Vec::new());
    };
    let root = parse_xml(comment_bytes)?;
    let locations = collect_comment_locations(package)?;
    let mut comments = Vec::new();
    for child in &root.children {
        let XMLNode::Element(element) = child else {
            continue;
        };
        if local_name(&element.name) != "comment" {
            continue;
        }
        let mut record = comment_record_from_element(element);
        record.locations = locations.get(&record.id).cloned().unwrap_or_default();
        comments.push(record);
    }
    Ok(comments)
}

fn collect_comment_locations(
    package: &DocxPackage,
) -> Result<std::collections::BTreeMap<u32, Vec<CommentLocation>>> {
    let mut output = std::collections::BTreeMap::<u32, Vec<CommentLocation>>::new();
    for part_name in comment_story_part_names(package) {
        let Some(bytes) = package.get_file(&part_name) else {
            continue;
        };
        let root = parse_xml(bytes)?;
        let story = story_container_ref(&root)?;
        for child in &story.children {
            let XMLNode::Element(paragraph) = child else {
                continue;
            };
            if local_name(&paragraph.name) != "p" {
                continue;
            }
            let mut ids = Vec::new();
            collect_comment_ids(paragraph, &mut ids);
            if ids.is_empty() {
                continue;
            }
            let text = paragraph_text(paragraph);
            for id in ids {
                output.entry(id).or_default().push(CommentLocation {
                    part: part_name.clone(),
                    paragraph_text: text.clone(),
                });
            }
        }
    }
    Ok(output)
}

fn comment_story_part_names(package: &DocxPackage) -> Vec<String> {
    package
        .entries
        .iter()
        .filter(|entry| {
            !entry.is_dir
                && (entry.name == DOCUMENT_XML_PATH
                    || entry.name.starts_with("word/header")
                    || entry.name.starts_with("word/footer"))
        })
        .map(|entry| entry.name.clone())
        .collect()
}

fn collect_comment_ids(element: &Element, output: &mut Vec<u32>) {
    if matches!(
        local_name(&element.name),
        "commentRangeStart" | "commentRangeEnd" | "commentReference"
    ) && let Some(id) = comment_id_of_element(element)
        && !output.contains(&id)
    {
        output.push(id);
    }

    for child in &element.children {
        if let XMLNode::Element(element) = child {
            collect_comment_ids(element, output);
        }
    }
}

fn element_contains_comment_markup(element: &Element) -> bool {
    matches!(
        local_name(&element.name),
        "commentRangeStart" | "commentRangeEnd" | "commentReference"
    ) || element.children.iter().any(|child| match child {
        XMLNode::Element(inner) => element_contains_comment_markup(inner),
        _ => false,
    })
}

fn comment_id_of_element(element: &Element) -> Option<u32> {
    element
        .attributes
        .get("w:id")
        .or_else(|| element.attributes.get("id"))
        .and_then(|value| value.parse::<u32>().ok())
}

fn comment_record_from_element(element: &Element) -> CommentRecord {
    let author = element
        .attributes
        .get("w:author")
        .or_else(|| element.attributes.get("author"))
        .cloned()
        .unwrap_or_default();
    let initials = element
        .attributes
        .get("w:initials")
        .or_else(|| element.attributes.get("initials"))
        .cloned()
        .filter(|value| !value.is_empty());
    let date = element
        .attributes
        .get("w:date")
        .or_else(|| element.attributes.get("date"))
        .cloned()
        .filter(|value| !value.is_empty());
    CommentRecord {
        id: comment_id_of_element(element).unwrap_or_default(),
        author,
        initials,
        date,
        highlight: first_highlight(element).unwrap_or_else(default_highlight_value),
        comment_text: paragraph_text(element),
        locations: Vec::new(),
    }
}

fn make_comment_entry_from_record(record: &CommentRecord) -> Result<Element> {
    let spec = CommentSpec {
        text: String::new(),
        comment_text: record.comment_text.clone(),
        author: record.author.clone(),
        initials: record.initials.clone(),
        date: record.date.clone(),
        style: ParagraphStyle::Normal,
        highlight: record.highlight.clone(),
    };
    make_comment_entry(&spec, record.id, "00000000")
}

fn first_highlight(element: &Element) -> Option<String> {
    if local_name(&element.name) == "highlight" {
        return element
            .attributes
            .get("w:val")
            .or_else(|| element.attributes.get("val"))
            .cloned();
    }
    for child in &element.children {
        if let XMLNode::Element(element) = child
            && let Some(value) = first_highlight(element)
        {
            return Some(value);
        }
    }
    None
}

fn remove_comment_markup_from_element(element: &mut Element, comment_id: u32) {
    let mut retained = Vec::new();
    for mut child in std::mem::take(&mut element.children) {
        let keep = match &mut child {
            XMLNode::Element(inner)
                if matches!(
                    local_name(&inner.name),
                    "commentRangeStart" | "commentRangeEnd" | "commentReference"
                ) && comment_id_of_element(inner) == Some(comment_id) =>
            {
                false
            }
            XMLNode::Element(inner) => {
                remove_comment_markup_from_element(inner, comment_id);
                !(local_name(&inner.name) == "r" && run_is_empty(inner))
            }
            _ => true,
        };
        if keep {
            retained.push(child);
        }
    }
    element.children = retained;
}

fn run_is_empty(run: &Element) -> bool {
    if local_name(&run.name) != "r" {
        return false;
    }
    !run.children.iter().any(|child| match child {
        XMLNode::Text(text) => !text.trim().is_empty(),
        XMLNode::Element(element) => !matches!(
            local_name(&element.name),
            "rPr" | "commentReference"
        ),
        _ => false,
    })
}

fn append_note_entry(
    package: &mut DocxPackage,
    kind: &NoteKind,
    note_id: u32,
    note: &NoteSpec,
) -> Result<()> {
    let part_name = match kind {
        NoteKind::Footnote => "word/footnotes.xml",
        NoteKind::Endnote => "word/endnotes.xml",
    };
    with_xml_part_mut(package, part_name, |root| {
        root.children
            .push(XMLNode::Element(make_note_entry(kind, note, note_id)?));
        Ok(())
    })
}

fn next_named_id(package: &DocxPackage, part_name: &str, element_local_name: &str) -> Result<u32> {
    let bytes = package
        .get_file(part_name)
        .with_context(|| format!("missing part {part_name}"))?;
    let root = parse_xml(bytes)?;
    let mut max_id = 0u32;
    collect_max_named_id(&root, element_local_name, &mut max_id);
    Ok(max_id + 1)
}

fn next_hex_attr_id(
    package: &DocxPackage,
    part_name: &str,
    element_local_name: &str,
    attr_local_name: &str,
) -> Result<String> {
    let bytes = package
        .get_file(part_name)
        .with_context(|| format!("missing part {part_name}"))?;
    let root = parse_xml(bytes)?;
    let mut max_id = 0u32;
    collect_max_hex_attr_id(&root, element_local_name, attr_local_name, &mut max_id);
    Ok(format!("{:08X}", max_id + 1))
}

fn collect_max_named_id(element: &Element, target_local_name: &str, max_id: &mut u32) {
    if local_name(&element.name) == target_local_name
        && let Some(id) = element
            .attributes
            .get("w:id")
            .or_else(|| element.attributes.get("id"))
        && let Ok(parsed) = id.parse::<i32>()
        && parsed >= 0
    {
        *max_id = (*max_id).max(parsed as u32);
    }

    for child in &element.children {
        if let XMLNode::Element(element) = child {
            collect_max_named_id(element, target_local_name, max_id);
        }
    }
}

fn collect_max_hex_attr_id(
    element: &Element,
    element_local_name: &str,
    attr_local_name: &str,
    max_id: &mut u32,
) {
    if local_name(&element.name) == element_local_name
        && let Some(value) = attr_value_local(element, attr_local_name)
        && let Ok(parsed) = u32::from_str_radix(value, 16)
    {
        *max_id = (*max_id).max(parsed);
    }

    for child in &element.children {
        if let XMLNode::Element(element) = child {
            collect_max_hex_attr_id(element, element_local_name, attr_local_name, max_id);
        }
    }
}

fn attr_value_local<'a>(element: &'a Element, attr_local_name: &str) -> Option<&'a str> {
    element.attributes.iter().find_map(|(key, value)| {
        key.rsplit(':')
            .next()
            .filter(|local| *local == attr_local_name)
            .map(|_| value.as_str())
    })
}

fn next_change_id(story: &Element) -> u32 {
    let mut max_id = 0u32;
    collect_max_change_id(story, &mut max_id);
    max_id + 1
}

fn collect_max_change_id(element: &Element, max_id: &mut u32) {
    if matches!(local_name(&element.name), "ins" | "del")
        && let Some(id) = element
            .attributes
            .get("w:id")
            .or_else(|| element.attributes.get("id"))
        && let Ok(parsed) = id.parse::<u32>()
    {
        *max_id = (*max_id).max(parsed);
    }

    for child in &element.children {
        if let XMLNode::Element(element) = child {
            collect_max_change_id(element, max_id);
        }
    }
}

fn next_sdt_id(story: &Element) -> u32 {
    let mut max_id = 0u32;
    collect_max_sdt_id(story, &mut max_id);
    max_id + 1
}

fn collect_max_sdt_id(element: &Element, max_id: &mut u32) {
    if local_name(&element.name) == "id"
        && let Some(id) = element
            .attributes
            .get("w:val")
            .or_else(|| element.attributes.get("val"))
        && let Ok(parsed) = id.parse::<u32>()
    {
        *max_id = (*max_id).max(parsed);
    }

    for child in &element.children {
        if let XMLNode::Element(element) = child {
            collect_max_sdt_id(element, max_id);
        }
    }
}

fn append_comment_compatibility_metadata(
    package: &mut DocxPackage,
    comment_para_id: &str,
    author: &str,
) -> Result<()> {
    if package.get_file("word/commentsExtended.xml").is_some() {
        with_xml_part_mut(package, "word/commentsExtended.xml", |root| {
            root.children.push(XMLNode::Element(make_comment_ex_element(comment_para_id)));
            Ok(())
        })?;
    }

    if package.get_file("word/commentsIds.xml").is_some() {
        let durable_id = next_hex_attr_id(package, "word/commentsIds.xml", "commentId", "durableId")?;
        with_xml_part_mut(package, "word/commentsIds.xml", |root| {
            root.children.push(XMLNode::Element(make_comment_id_element(
                comment_para_id,
                &durable_id,
            )));
            Ok(())
        })?;
    }

    if package.get_file("word/people.xml").is_some() {
        let already_exists = {
            let bytes = package
                .get_file("word/people.xml")
                .with_context(|| "missing part word/people.xml".to_string())?;
            let root = parse_xml(bytes)?;
            root.children.iter().any(|child| match child {
                XMLNode::Element(element) if local_name(&element.name) == "person" => {
                    attr_value_local(element, "author") == Some(author)
                }
                _ => false,
            })
        };
        if !already_exists {
            with_xml_part_mut(package, "word/people.xml", |root| {
                root.children
                    .push(XMLNode::Element(make_person_element(author)));
                Ok(())
            })?;
        }
    }

    Ok(())
}

fn simple_element(name: &str) -> Element {
    Element::new(name)
}

fn simple_text_element(name: &str, text: &str) -> Element {
    let mut element = simple_element(name);
    element.children.push(XMLNode::Text(text.to_string()));
    element
}

fn make_comment_ex_element(comment_para_id: &str) -> Element {
    let mut element = simple_element("w15:commentEx");
    element
        .attributes
        .insert("w15:paraId".to_string(), comment_para_id.to_string());
    element
        .attributes
        .insert("w15:done".to_string(), "0".to_string());
    element
}

fn make_comment_id_element(comment_para_id: &str, durable_id: &str) -> Element {
    let mut element = simple_element("w16cid:commentId");
    element
        .attributes
        .insert("w16cid:paraId".to_string(), comment_para_id.to_string());
    element
        .attributes
        .insert("w16cid:durableId".to_string(), durable_id.to_string());
    element
}

fn make_person_element(author: &str) -> Element {
    let mut element = simple_element("w15:person");
    element
        .attributes
        .insert("w15:author".to_string(), author.to_string());
    element
}

fn make_comment_entry(comment: &CommentSpec, comment_id: u32, comment_para_id: &str) -> Result<Element> {
    let initials = comment.initials.as_deref().unwrap_or("");
    let date = comment.date.as_deref().unwrap_or("2026-01-01T00:00:00Z");

    let mut comment_element = simple_element("w:comment");
    comment_element
        .attributes
        .insert("w:id".to_string(), comment_id.to_string());
    comment_element
        .attributes
        .insert("w:author".to_string(), comment.author.clone());
    comment_element
        .attributes
        .insert("w:initials".to_string(), initials.to_string());
    comment_element
        .attributes
        .insert("w:date".to_string(), date.to_string());

    let mut paragraph = simple_element("w:p");
    paragraph
        .attributes
        .insert("w15:paraId".to_string(), comment_para_id.to_string());

    let mut ppr = simple_element("w:pPr");
    let mut pstyle = simple_element("w:pStyle");
    pstyle
        .attributes
        .insert("w:val".to_string(), "CommentText".to_string());
    ppr.children.push(XMLNode::Element(pstyle));
    paragraph.children.push(XMLNode::Element(ppr));

    let mut annotation_run = simple_element("w:r");
    let mut annotation_rpr = simple_element("w:rPr");
    let mut annotation_style = simple_element("w:rStyle");
    annotation_style
        .attributes
        .insert("w:val".to_string(), "CommentReference".to_string());
    annotation_rpr.children.push(XMLNode::Element(annotation_style));
    annotation_run
        .children
        .push(XMLNode::Element(annotation_rpr));
    annotation_run
        .children
        .push(XMLNode::Element(simple_element("w:annotationRef")));
    paragraph.children.push(XMLNode::Element(annotation_run));

    let mut text_run = simple_element("w:r");
    let mut text_rpr = simple_element("w:rPr");
    let mut highlight = simple_element("w:highlight");
    highlight
        .attributes
        .insert("w:val".to_string(), comment.highlight.clone());
    text_rpr.children.push(XMLNode::Element(highlight));
    text_run.children.push(XMLNode::Element(text_rpr));

    let mut text = simple_text_element("w:t", &comment.comment_text);
    text.attributes
        .insert("xml:space".to_string(), "preserve".to_string());
    text_run.children.push(XMLNode::Element(text));
    paragraph.children.push(XMLNode::Element(text_run));

    comment_element.children.push(XMLNode::Element(paragraph));
    Ok(comment_element)
}

fn attach_comment_to_paragraph(paragraph: &mut Element, comment_id: u32) -> Result<()> {
    if element_contains_comment_markup(paragraph) {
        bail!(
            "target paragraph already contains comment markup; add-comment cannot safely add another comment to the same paragraph yet"
        );
    }

    let mut start = simple_element("w:commentRangeStart");
    start
        .attributes
        .insert("w:id".to_string(), comment_id.to_string());

    let mut end = simple_element("w:commentRangeEnd");
    end.attributes
        .insert("w:id".to_string(), comment_id.to_string());

    let mut reference = simple_element("w:r");
    let mut reference_rpr = simple_element("w:rPr");
    let mut reference_style = simple_element("w:rStyle");
    reference_style
        .attributes
        .insert("w:val".to_string(), "CommentReference".to_string());
    reference_rpr
        .children
        .push(XMLNode::Element(reference_style));
    reference.children.push(XMLNode::Element(reference_rpr));

    let mut comment_reference = simple_element("w:commentReference");
    comment_reference
        .attributes
        .insert("w:id".to_string(), comment_id.to_string());
    reference
        .children
        .push(XMLNode::Element(comment_reference));

    let insert_at = paragraph
        .children
        .iter()
        .position(|child| !matches!(child, XMLNode::Element(element) if local_name(&element.name) == "pPr"))
        .unwrap_or(paragraph.children.len());
    paragraph.children.insert(insert_at, XMLNode::Element(start));
    paragraph.children.push(XMLNode::Element(end));
    paragraph.children.push(XMLNode::Element(reference));
    Ok(())
}

fn make_note_entry(kind: &NoteKind, note: &NoteSpec, note_id: u32) -> Result<Element> {
    let element_name = match kind {
        NoteKind::Footnote => "footnote",
        NoteKind::Endnote => "endnote",
    };
    parse_wrapped_fragment(&format!(
        r#"<w:{element_name} w:id="{note_id}">
<w:p><w:r><w:rPr><w:highlight w:val="{highlight}" /></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>
</w:{element_name}>"#,
        element_name = element_name,
        note_id = note_id,
        highlight = escape_xml_attr(&note.highlight),
        text = escape_xml_text(&note.body),
    ))
}

fn make_note_reference_paragraph(
    kind: &NoteKind,
    note: &NoteSpec,
    note_id: u32,
) -> Result<Element> {
    let reference_element = match kind {
        NoteKind::Footnote => "footnoteReference",
        NoteKind::Endnote => "endnoteReference",
    };
    parse_wrapped_fragment(&format!(
        r#"<w:p>
<w:pPr><w:pStyle w:val="{style}" /></w:pPr>
<w:r><w:rPr><w:highlight w:val="{highlight}" /></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>
<w:r><w:{reference_element} w:id="{note_id}" /></w:r>
</w:p>"#,
        style = note.style.style_id(),
        highlight = escape_xml_attr(&note.highlight),
        text = escape_xml_text(&note.reference_text),
        reference_element = reference_element,
        note_id = note_id,
    ))
}

fn make_content_control(control: &ContentControlSpec, control_id: u32) -> Result<Element> {
    let alias = control.alias.as_deref().unwrap_or(&control.tag);
    let placeholder = control.placeholder.as_deref().unwrap_or("");
    let placeholder_xml = if placeholder.is_empty() {
        String::new()
    } else {
        format!(
            r#"<w:placeholder><w:docPart w:val="{}" /></w:placeholder>"#,
            escape_xml_attr(placeholder)
        )
    };
    parse_wrapped_fragment(&format!(
        r#"<w:sdt>
<w:sdtPr>
<w:id w:val="{control_id}" />
<w:tag w:val="{tag}" />
<w:alias w:val="{alias}" />
{placeholder_xml}
</w:sdtPr>
<w:sdtContent>
<w:p>
<w:pPr><w:pStyle w:val="{style}" /></w:pPr>
<w:r><w:rPr><w:highlight w:val="{highlight}" /></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>
</w:p>
</w:sdtContent>
</w:sdt>"#,
        control_id = control_id,
        tag = escape_xml_attr(&control.tag),
        alias = escape_xml_attr(alias),
        placeholder_xml = placeholder_xml,
        style = control.style.style_id(),
        highlight = escape_xml_attr(&control.highlight),
        text = escape_xml_text(&control.text),
    ))
}

fn make_field_paragraph(field: &FieldSpec) -> Result<Element> {
    parse_wrapped_fragment(&format!(
        r#"<w:p>
<w:pPr><w:pStyle w:val="{style}" /></w:pPr>
<w:r><w:fldChar w:fldCharType="begin" /></w:r>
<w:r><w:instrText xml:space="preserve">{instruction}</w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate" /></w:r>
<w:r><w:rPr><w:highlight w:val="{highlight}" /></w:rPr><w:t xml:space="preserve">{result}</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end" /></w:r>
</w:p>"#,
        style = field.style.style_id(),
        instruction = escape_xml_text(&field.instruction),
        highlight = escape_xml_attr(&field.highlight),
        result = escape_xml_text(&field.result),
    ))
}

fn make_tracked_insert_paragraph(
    entry: &ParagraphEntry,
    change_id: u32,
    author: &str,
    date: &str,
) -> Result<Element> {
    parse_wrapped_fragment(&format!(
        r#"<w:p>
<w:pPr><w:pStyle w:val="{style}" /></w:pPr>
<w:ins w:id="{change_id}" w:author="{author}" w:date="{date}">
<w:r><w:rPr>{run_props}<w:highlight w:val="{highlight}" /></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>
</w:ins>
</w:p>"#,
        style = entry.style.style_id(),
        change_id = change_id,
        author = escape_xml_attr(author),
        date = escape_xml_attr(date),
        run_props = run_props_xml(entry.bold, entry.italic, entry.underline),
        highlight = escape_xml_attr(&entry.highlight),
        text = escape_xml_text(&entry.text),
    ))
}

fn make_tracked_delete_paragraph(
    text: &str,
    change_id: u32,
    author: &str,
    date: &str,
) -> Result<Element> {
    parse_wrapped_fragment(&format!(
        r#"<w:p>
<w:del w:id="{change_id}" w:author="{author}" w:date="{date}">
<w:r><w:delText xml:space="preserve">{text}</w:delText></w:r>
</w:del>
</w:p>"#,
        change_id = change_id,
        author = escape_xml_attr(author),
        date = escape_xml_attr(date),
        text = escape_xml_text(text),
    ))
}

fn minimal_comments_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>"#
        .to_vec()
}

fn minimal_comments_extended_xml() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="{W15_NS}"></w15:commentsEx>"#
    )
    .into_bytes()
}

fn minimal_comments_ids_xml() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w16cid:commentsIds xmlns:w16cid="{W16CID_NS}"></w16cid:commentsIds>"#
    )
    .into_bytes()
}

fn minimal_people_xml() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:people xmlns:w15="{W15_NS}"></w15:people>"#
    )
    .into_bytes()
}

fn minimal_footnotes_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator /></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator /></w:r></w:p></w:footnote>
</w:footnotes>"#
        .to_vec()
}

fn minimal_endnotes_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator /></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator /></w:r></w:p></w:endnote>
</w:endnotes>"#
        .to_vec()
}

fn minimal_header_footer_xml(root_name: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><{root_name} xmlns:w="{W_NS}"><w:p /></{root_name}>"#
    )
    .into_bytes()
}

fn make_text_paragraph(entry: &ParagraphEntry) -> Result<Element> {
    parse_wrapped_fragment(&format!(
        r#"<w:p>
<w:pPr><w:pStyle w:val="{style}" /></w:pPr>
<w:r>
<w:rPr>{run_props}<w:highlight w:val="{highlight}" /></w:rPr>
<w:t xml:space="preserve">{text}</w:t>
</w:r>
</w:p>"#,
        style = entry.style.style_id(),
        highlight = escape_xml_attr(&entry.highlight),
        text = escape_xml_text(&entry.text),
        run_props = run_props_xml(entry.bold, entry.italic, entry.underline),
    ))
}

fn raw_text_paragraph_xml(entry: &ParagraphEntry) -> String {
    format!(
        r#"<w:p><w:pPr><w:pStyle w:val="{style}" /></w:pPr><w:r><w:rPr>{run_props}<w:highlight w:val="{highlight}" /></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#,
        style = entry.style.style_id(),
        run_props = raw_run_props_xml(entry.bold, entry.italic, entry.underline),
        highlight = escape_xml_attr(&entry.highlight),
        text = escape_xml_text(&entry.text),
    )
}

fn raw_image_paragraph_xml(image: &ImageSpec, relationship_id: &str, drawing_id: u32) -> String {
    let alt_text = image.alt_text.as_deref().unwrap_or("Inserted image");
    let pic_name = format!("Picture {drawing_id}");
    format!(
        r#"<w:p><w:r><w:rPr><w:highlight w:val="{highlight}" /></w:rPr><w:drawing><wp:inline xmlns:wp="{wp_ns}" xmlns:a="{a_ns}" xmlns:pic="{pic_ns}" xmlns:r="{r_ns}"><wp:extent cx="{cx}" cy="{cy}" /><wp:docPr id="{drawing_id}" name="{pic_name}" descr="{alt_text}" /><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1" /></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="{pic_ns}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="{pic_name}" /><pic:cNvPicPr /></pic:nvPicPr><pic:blipFill><a:blip r:embed="{relationship_id}" /><a:stretch><a:fillRect /></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="{cx}" cy="{cy}" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
        highlight = DEFAULT_HIGHLIGHT,
        wp_ns = WP_NS,
        a_ns = A_NS,
        pic_ns = PIC_NS,
        r_ns = R_NS,
        cx = image.width_emu,
        cy = image.height_emu,
        drawing_id = drawing_id,
        pic_name = escape_xml_attr(&pic_name),
        alt_text = escape_xml_attr(alt_text),
        relationship_id = escape_xml_attr(relationship_id),
    )
}

fn raw_run_props_xml(bold: bool, italic: bool, underline: bool) -> String {
    let mut output = String::new();
    if bold {
        output.push_str("<w:b />");
    }
    if italic {
        output.push_str("<w:i />");
    }
    if underline {
        output.push_str(r#"<w:u w:val="single" />"#);
    }
    output
}

fn make_hyperlink_paragraph(hyperlink: &HyperlinkSpec, relationship_id: &str) -> Result<Element> {
    parse_wrapped_fragment(&format!(
        r#"<w:p>
<w:pPr><w:pStyle w:val="{style}" /></w:pPr>
<w:hyperlink r:id="{relationship_id}" w:history="1">
<w:r>
<w:rPr><w:rStyle w:val="Hyperlink" /><w:highlight w:val="{highlight}" /></w:rPr>
<w:t xml:space="preserve">{text}</w:t>
</w:r>
</w:hyperlink>
</w:p>"#,
        style = hyperlink.style.style_id(),
        relationship_id = escape_xml_attr(relationship_id),
        highlight = escape_xml_attr(&hyperlink.highlight),
        text = escape_xml_text(&hyperlink.text),
    ))
}

fn make_table(table: &TableSpec) -> Result<Element> {
    let mut rows_xml = String::new();
    for row in &table.rows {
        rows_xml.push_str("<w:tr>");
        for cell in &row.cells {
            rows_xml.push_str(&format!(
                r#"<w:tc><w:p><w:r><w:rPr><w:highlight w:val="{highlight}" /></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>"#,
                highlight = escape_xml_attr(&table.highlight),
                text = escape_xml_text(cell),
            ));
        }
        rows_xml.push_str("</w:tr>");
    }

    let table_props = if let Some(style) = &table.style {
        format!(
            r#"<w:tblPr><w:tblStyle w:val="{}" /></w:tblPr>"#,
            escape_xml_attr(style)
        )
    } else {
        "<w:tblPr />".to_string()
    };

    parse_wrapped_fragment(&format!(
        r#"<w:tbl>{table_props}<w:tblGrid />{rows_xml}</w:tbl>"#
    ))
}

fn make_section_break_paragraph(break_type: &SectionBreakType) -> Result<Element> {
    parse_wrapped_fragment(&format!(
        r#"<w:p><w:pPr><w:sectPr><w:type w:val="{}" /></w:sectPr></w:pPr></w:p>"#,
        break_type.word_value()
    ))
}

fn make_image_paragraph(
    image: &ImageSpec,
    relationship_id: &str,
    drawing_id: u32,
) -> Result<Element> {
    let alt_text = image.alt_text.as_deref().unwrap_or("Inserted image");
    let pic_name = format!("Picture {drawing_id}");
    parse_wrapped_fragment(&format!(
        r#"<w:p>
<w:r>
<w:rPr><w:highlight w:val="{highlight}" /></w:rPr>
<w:drawing>
<wp:inline xmlns:wp="{wp_ns}" xmlns:a="{a_ns}" xmlns:pic="{pic_ns}" xmlns:r="{r_ns}">
<wp:extent cx="{cx}" cy="{cy}" />
<wp:docPr id="{drawing_id}" name="{pic_name}" descr="{alt_text}" />
<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1" /></wp:cNvGraphicFramePr>
<a:graphic>
<a:graphicData uri="{pic_ns}">
<pic:pic>
<pic:nvPicPr>
<pic:cNvPr id="0" name="{pic_name}" />
<pic:cNvPicPr />
</pic:nvPicPr>
<pic:blipFill>
<a:blip r:embed="{relationship_id}" />
<a:stretch><a:fillRect /></a:stretch>
</pic:blipFill>
<pic:spPr>
<a:xfrm><a:off x="0" y="0" /><a:ext cx="{cx}" cy="{cy}" /></a:xfrm>
<a:prstGeom prst="rect"><a:avLst /></a:prstGeom>
</pic:spPr>
</pic:pic>
</a:graphicData>
</a:graphic>
</wp:inline>
</w:drawing>
</w:r>
</w:p>"#,
        highlight = DEFAULT_HIGHLIGHT,
        wp_ns = WP_NS,
        a_ns = A_NS,
        pic_ns = PIC_NS,
        r_ns = R_NS,
        cx = image.width_emu,
        cy = image.height_emu,
        drawing_id = drawing_id,
        pic_name = escape_xml_attr(&pic_name),
        alt_text = escape_xml_attr(alt_text),
        relationship_id = escape_xml_attr(relationship_id),
    ))
}

fn run_props_xml(bold: bool, italic: bool, underline: bool) -> String {
    let mut output = String::new();
    if bold {
        output.push_str("<w:b />");
    }
    if italic {
        output.push_str("<w:i />");
    }
    if underline {
        output.push_str(r#"<w:u w:val="single" />"#);
    }
    output
}

fn parse_wrapped_fragment(fragment: &str) -> Result<Element> {
    let wrapped = format!(
        r#"<root xmlns:w="{W_NS}" xmlns:w15="{W15_NS}" xmlns:w16cid="{W16CID_NS}" xmlns:r="{R_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}">{fragment}</root>"#
    );
    let root = parse_xml(wrapped.as_bytes())?;
    root.children
        .into_iter()
        .find_map(|child| match child {
            XMLNode::Element(element) => Some(element),
            _ => None,
        })
        .context("fragment did not produce an element")
}

fn parse_xml(bytes: &[u8]) -> Result<Element> {
    Element::parse(Cursor::new(bytes)).context("failed to parse XML")
}

fn write_xml(element: &Element) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    element
        .write_with_config(
            &mut out,
            EmitterConfig::new()
                .perform_indent(false)
                .write_document_declaration(true),
        )
        .context("failed to serialize XML")?;
    Ok(out)
}

fn find_child_mut_local<'a>(element: &'a mut Element, local: &str) -> Option<&'a mut Element> {
    for child in &mut element.children {
        if let XMLNode::Element(element) = child
            && local_name(&element.name) == local
        {
            return Some(element);
        }
    }
    None
}

fn find_child_ref_local<'a>(element: &'a Element, local: &str) -> Option<&'a Element> {
    for child in &element.children {
        if let XMLNode::Element(element) = child
            && local_name(&element.name) == local
        {
            return Some(element);
        }
    }
    None
}

fn local_name(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

fn escape_xml_text(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn escape_xml_attr(value: &str) -> String {
    escape_xml_text(value)
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn decode_xml_entities(value: &str) -> String {
    value
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

fn minimal_part_relationships_xml() -> Vec<u8> {
    format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELS_NS}"></Relationships>"#)
        .into_bytes()
}

fn minimal_root_relationships_xml() -> Vec<u8> {
    minimal_part_relationships_xml()
}

fn minimal_core_properties_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>"#
        .to_vec()
}

#[derive(Debug, Clone)]
struct PackageEntry {
    name: String,
    bytes: Vec<u8>,
    is_dir: bool,
    compression: CompressionMethod,
}

#[derive(Debug, Clone)]
struct DocxPackage {
    entries: Vec<PackageEntry>,
}

impl DocxPackage {
    fn read(input_bytes: &[u8]) -> Result<Self> {
        let cursor = Cursor::new(input_bytes.to_vec());
        let mut archive = ZipArchive::new(cursor).context("failed to open .docx zip package")?;
        let mut entries = Vec::new();

        for index in 0..archive.len() {
            let mut file = archive
                .by_index(index)
                .context("failed to read zip entry")?;
            let mut bytes = Vec::new();
            if !file.is_dir() {
                file.read_to_end(&mut bytes)
                    .with_context(|| format!("failed to read zip entry {}", file.name()))?;
            }
            entries.push(PackageEntry {
                name: file.name().to_string(),
                bytes,
                is_dir: file.is_dir(),
                compression: file.compression(),
            });
        }

        Ok(Self { entries })
    }

    fn write(&self) -> Result<Vec<u8>> {
        let mut output = Cursor::new(Vec::new());
        {
            let mut writer = ZipWriter::new(&mut output);
            for entry in &self.entries {
                let options = SimpleFileOptions::default()
                    .compression_method(entry.compression)
                    .unix_permissions(if entry.is_dir { 0o755 } else { 0o644 });
                if entry.is_dir {
                    writer
                        .add_directory(&entry.name, options)
                        .with_context(|| format!("failed to write directory {}", entry.name))?;
                } else {
                    writer
                        .start_file(&entry.name, options)
                        .with_context(|| format!("failed to start file {}", entry.name))?;
                    writer
                        .write_all(&entry.bytes)
                        .with_context(|| format!("failed to write file {}", entry.name))?;
                }
            }
            writer
                .finish()
                .context("failed to finalize output package")?;
        }
        Ok(output.into_inner())
    }

    fn get_file(&self, name: &str) -> Option<&[u8]> {
        self.entries
            .iter()
            .find(|entry| !entry.is_dir && entry.name == name)
            .map(|entry| entry.bytes.as_slice())
    }

    fn set_file(&mut self, name: &str, bytes: Vec<u8>) {
        if let Some(entry) = self
            .entries
            .iter_mut()
            .find(|entry| !entry.is_dir && entry.name == name)
        {
            entry.bytes = bytes;
            entry.compression = CompressionMethod::Deflated;
            return;
        }

        self.entries.push(PackageEntry {
            name: name.to_string(),
            bytes,
            is_dir: false,
            compression: CompressionMethod::Deflated,
        });
    }

    fn ensure_directory(&mut self, name: &str) {
        let normalized = if name.ends_with('/') {
            name.to_string()
        } else {
            format!("{name}/")
        };
        if self
            .entries
            .iter()
            .any(|entry| entry.is_dir && entry.name == normalized)
        {
            return;
        }
        self.entries.push(PackageEntry {
            name: normalized,
            bytes: Vec::new(),
            is_dir: true,
            compression: CompressionMethod::Stored,
        });
    }

    fn file_map(&self) -> std::collections::BTreeMap<String, Vec<u8>> {
        self.entries
            .iter()
            .filter(|entry| !entry.is_dir)
            .map(|entry| (entry.name.clone(), entry.bytes.clone()))
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    const MINIMAL_DOCUMENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Purpose</w:t></w:r></w:p>
    <w:p><w:r><w:t>Strategic principles</w:t></w:r></w:p>
    <w:p><w:r><w:t>future expansion opportunity without diluting the current priority</w:t></w:r></w:p>
    <w:sectPr />
  </w:body>
</w:document>"#;

    #[test]
    fn supports_practical_document_operations() {
        let spec = AutomationSpec {
            operations: vec![
                Operation::InsertParagraphs {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("future expansion opportunity".to_string()),
                    entries: vec![ParagraphEntry {
                        text: "GitHub and VS Code".to_string(),
                        style: ParagraphStyle::ListBullet,
                        highlight: "green".to_string(),
                        bold: false,
                        italic: false,
                        underline: false,
                    }],
                },
                Operation::InsertTableAfter {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                    table: TableSpec {
                        rows: vec![TableRowSpec {
                            cells: vec!["Metric".to_string(), "Value".to_string()],
                        }],
                        style: Some("TableGrid".to_string()),
                        highlight: "green".to_string(),
                    },
                },
                Operation::ReplaceText {
                    part: PartTarget::default(),
                    find: "Purpose".to_string(),
                    replace: "Updated Purpose".to_string(),
                    highlight: "green".to_string(),
                },
                Operation::SetCoreProperty {
                    property: CoreProperty::Title,
                    value: "Project Strategy".to_string(),
                },
            ],
        };

        let output =
            apply_spec_to_docx_bytes(&make_minimal_docx().expect("docx"), &spec, Path::new("."))
                .expect("operations should succeed");
        let mut archive = ZipArchive::new(Cursor::new(output)).expect("zip output");

        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document xml")
            .read_to_string(&mut document)
            .expect("read document");
        assert!(document.contains("Updated Purpose"));
        assert!(document.contains("GitHub and VS Code"));
        assert!(document.contains("tblStyle"));
        assert!(document.contains("TableGrid"));
        assert!(document.contains(r#"<w:highlight w:val="green" />"#));

        let mut core = String::new();
        archive
            .by_name("docProps/core.xml")
            .expect("core xml")
            .read_to_string(&mut core)
            .expect("read core");
        assert!(core.contains("Project Strategy"));
    }

    #[test]
    fn raw_insert_path_preserves_namespaced_attributes() {
        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("future expansion opportunity".to_string()),
                entries: vec![ParagraphEntry {
                    text: "GitHub and VS Code".to_string(),
                    style: ParagraphStyle::ListBullet,
                    highlight: "green".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let output =
            apply_spec_to_docx_bytes(&make_minimal_docx().expect("docx"), &spec, Path::new("."))
                .expect("raw insert path should succeed");
        let mut archive = ZipArchive::new(Cursor::new(output)).expect("zip output");
        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document xml")
            .read_to_string(&mut document)
            .expect("read document");

        assert!(document.contains(r#"<w:pStyle w:val="ListBullet""#));
        assert!(document.contains(r#"xml:space="preserve""#));
        assert!(!document.contains(r#"<w:pStyle val="ListBullet""#));
        assert!(!document.contains(r#"<w:t space="preserve""#));
    }

    #[test]
    fn creates_relationships_for_hyperlinks_and_images() {
        let dir = tempdir().expect("temp dir");
        let image_path = dir.path().join("sample.png");
        fs::write(&image_path, [0x89, b'P', b'N', b'G']).expect("write image");

        let spec = AutomationSpec {
            operations: vec![
                Operation::InsertHyperlinkAfter {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                    hyperlink: HyperlinkSpec {
                        text: "Open GitHub".to_string(),
                        url: "https://github.com".to_string(),
                        style: ParagraphStyle::Normal,
                        highlight: "green".to_string(),
                    },
                },
                Operation::InsertImageAfter {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                    image: ImageSpec {
                        path: image_path.to_string_lossy().to_string(),
                        width_emu: 914400,
                        height_emu: 914400,
                        alt_text: Some("Sample".to_string()),
                    },
                },
            ],
        };

        let output =
            apply_spec_to_docx_bytes(&make_minimal_docx().expect("docx"), &spec, Path::new(""))
                .expect("operations should succeed");
        let mut archive = ZipArchive::new(Cursor::new(output)).expect("zip output");

        let mut rels = String::new();
        archive
            .by_name("word/_rels/document.xml.rels")
            .expect("document rels")
            .read_to_string(&mut rels)
            .expect("read rels");
        assert!(rels.contains("https://github.com"));
        assert!(rels.contains(REL_TYPE_IMAGE));

        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document")
            .read_to_string(&mut document)
            .expect("document string");
        assert!(document.contains("Open GitHub"));
        assert!(document.contains("<w:drawing>"));

        let mut content_types = String::new();
        archive
            .by_name(CONTENT_TYPES_XML_PATH)
            .expect("content types")
            .read_to_string(&mut content_types)
            .expect("read content types");
        assert!(content_types.contains(r#"Extension="png""#));
        assert!(archive.by_name("word/media/image1.png").is_ok());
    }

    #[test]
    fn raw_image_insert_path_preserves_namespaced_attributes() {
        let dir = tempdir().expect("temp dir");
        let image_path = dir.path().join("sample.gif");
        fs::write(&image_path, b"GIF89a").expect("write gif");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertImageAfter {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                image: ImageSpec {
                    path: image_path.to_string_lossy().to_string(),
                    width_emu: 914400,
                    height_emu: 914400,
                    alt_text: Some("Animated GIF sample".to_string()),
                },
            }],
        };

        let output = apply_spec_to_docx_bytes(&make_word_like_docx().expect("docx"), &spec, Path::new(""))
            .expect("image insert should succeed");
        let mut archive = ZipArchive::new(Cursor::new(output)).expect("zip output");
        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document xml")
            .read_to_string(&mut document)
            .expect("read document");

        assert!(document.contains(r#"mc:Ignorable="w14 wp14""#));
        assert!(document.contains(r#"<w:highlight w:val="green" />"#));
        assert!(document.contains(r#"<a:blip r:embed="rId"#));
        assert!(!document.contains(r#"<w:highlight val="green" />"#));
        assert!(!document.contains(r#"<a:blip embed="rId"#));

        let mut content_types = String::new();
        archive
            .by_name(CONTENT_TYPES_XML_PATH)
            .expect("content types")
            .read_to_string(&mut content_types)
            .expect("read content types");
        assert!(content_types.contains(r#"Extension="gif""#));
        assert!(archive.by_name("word/media/image1.gif").is_ok());
    }

    #[test]
    fn supports_second_phase_word_operations() {
        let spec = AutomationSpec {
            operations: vec![
                Operation::InsertCommentAfter {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                    comment: CommentSpec {
                        text: "Scenario with reviewer context".to_string(),
                        comment_text: "Add the workflow rationale.".to_string(),
                        author: "Copilot".to_string(),
                        initials: Some("CP".to_string()),
                        date: Some("2026-05-18T00:00:00Z".to_string()),
                        style: ParagraphStyle::Normal,
                        highlight: "green".to_string(),
                    },
                },
                Operation::InsertNoteAfter {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                    kind: NoteKind::Footnote,
                    note: NoteSpec {
                        reference_text: "Automated review and publishing steps".to_string(),
                        body: "This note captures a reusable workflow pattern."
                            .to_string(),
                        style: ParagraphStyle::Normal,
                        highlight: "green".to_string(),
                    },
                },
                Operation::InsertContentControlAfter {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                    control: ContentControlSpec {
                        tag: "scenario-summary".to_string(),
                        alias: Some("Scenario Summary".to_string()),
                        text: "Tracked scenario summary".to_string(),
                        placeholder: Some("DefaultPlaceholder".to_string()),
                        style: ParagraphStyle::Quote,
                        highlight: "green".to_string(),
                    },
                },
                Operation::InsertFieldAfter {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                    field: FieldSpec {
                        instruction: " PAGE ".to_string(),
                        result: "1".to_string(),
                        style: ParagraphStyle::Normal,
                        highlight: "green".to_string(),
                    },
                },
                Operation::TrackInsertParagraphs {
                    part: PartTarget::default(),
                    anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                    author: "Copilot".to_string(),
                    date: "2026-05-18T00:00:00Z".to_string(),
                    entries: vec![ParagraphEntry {
                        text: "Tracked inserted paragraph".to_string(),
                        style: ParagraphStyle::Normal,
                        highlight: "green".to_string(),
                        bold: false,
                        italic: false,
                        underline: false,
                    }],
                },
                Operation::TrackDeleteParagraphs {
                    part: PartTarget::default(),
                    contains: "Purpose".to_string(),
                    author: "Copilot".to_string(),
                    date: "2026-05-18T00:00:00Z".to_string(),
                },
            ],
        };

        let output =
            apply_spec_to_docx_bytes(&make_minimal_docx().expect("docx"), &spec, Path::new("."))
                .expect("operations should succeed");
        let mut archive = ZipArchive::new(Cursor::new(output)).expect("zip output");

        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document xml")
            .read_to_string(&mut document)
            .expect("read document");
        assert!(document.contains("commentRangeStart"));
        assert!(document.contains("footnoteReference"));
        assert!(document.contains("w:sdt"));
        assert!(document.contains("fldCharType=\"begin\""));
        assert!(document.contains("Tracked inserted paragraph"));
        assert!(document.contains("delText"));

        let mut comments = String::new();
        archive
            .by_name("word/comments.xml")
            .expect("comments part")
            .read_to_string(&mut comments)
            .expect("read comments");
        assert!(comments.contains("Add the workflow rationale."));

        let mut footnotes = String::new();
        archive
            .by_name("word/footnotes.xml")
            .expect("footnotes part")
            .read_to_string(&mut footnotes)
            .expect("read footnotes");
        assert!(footnotes.contains("reusable workflow pattern"));

        let mut rels = String::new();
        archive
            .by_name("word/_rels/document.xml.rels")
            .expect("document rels")
            .read_to_string(&mut rels)
            .expect("read rels");
        assert!(rels.contains(REL_TYPE_COMMENTS));
        assert!(rels.contains(REL_TYPE_FOOTNOTES));
    }

    #[test]
    fn supports_header_footer_upsert_and_anchor_queries() {
        let spec = AutomationSpec {
            operations: vec![Operation::UpsertHeaderFooter {
                kind: HeaderFooterKind::Header,
                reference: HeaderFooterReferenceKind::Default,
                section_index: 0,
                entries: vec![ParagraphEntry {
                    text: "Project Strategy Header".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "green".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let output =
            apply_spec_to_docx_bytes(&make_minimal_docx().expect("docx"), &spec, Path::new("."))
                .expect("upsert header");
        let report = validate_docx_bytes(&output).expect("validate output");
        assert!(report.is_valid(), "{:?}", report.issues);

        let mut archive = ZipArchive::new(Cursor::new(output.clone())).expect("zip output");
        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document xml")
            .read_to_string(&mut document)
            .expect("read document");
        assert!(document.contains("headerReference"));
        assert_eq!(document.matches("xmlns:r=").count(), 1);

        let mut header = String::new();
        archive
            .by_name("word/header1.xml")
            .expect("header part")
            .read_to_string(&mut header)
            .expect("read header");
        assert!(header.contains("Project Strategy Header"));

        let input_path = tempdir().expect("temp dir");
        let source = input_path.path().join("source.docx");
        let updated = input_path.path().join("updated.docx");
        fs::write(&source, make_repeated_anchor_docx().expect("source docx"))
            .expect("write source");
        fs::write(&updated, output).expect("write updated");

        let anchors = find_anchors_in_docx(
            &source,
            Some(DOCUMENT_XML_PATH),
            &AnchorTarget::Structured(AnchorSpec {
                text: "Strategic principles".to_string(),
                mode: AnchorMatchMode::Equals,
                occurrence: 2,
            }),
        )
        .expect("find anchors");
        assert_eq!(anchors.len(), 1);
        assert_eq!(anchors[0].index, 2);

        let diff = diff_docx_files(&source, &updated).expect("diff docs");
        assert!(
            diff.added_parts
                .iter()
                .any(|part| part == "word/header1.xml")
        );
        assert!(
            diff.changed_parts
                .iter()
                .any(|part| part == DOCUMENT_XML_PATH)
        );
    }

    #[test]
    fn applies_spec_from_json_file() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("input.docx");
        let output_path = dir.path().join("output.docx");
        let spec_path = dir.path().join("spec.json");

        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");
        fs::write(
            &spec_path,
            r#"{
  "operations": [
    {
      "type": "insert-section-break-after",
      "anchor": "Strategic principles",
      "break_type": "continuous"
    },
    {
      "type": "delete-paragraphs",
      "contains": "Purpose"
    },
    {
      "type": "insert-field-after",
      "anchor": "Strategic principles",
      "field": {
        "instruction": " PAGE ",
        "result": "1"
      }
    }
  ]
}"#,
        )
        .expect("write spec");

        let spec = AutomationSpec::from_path(&spec_path).expect("spec parse");
        apply_spec_file_to_docx(&input_path, &output_path, &spec, dir.path()).expect("apply");

        let bytes = fs::read(&output_path).expect("read output");
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip");
        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document")
            .read_to_string(&mut document)
            .expect("string");

        assert!(!document.contains(">Purpose<"));
        assert!(document.contains("continuous"));
        assert!(document.contains("fldCharType"));
    }

    #[test]
    fn publish_workflow_uses_temp_candidate_and_cleans_up_on_success() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("input.docx");
        let output_path = dir.path().join("published.docx");
        let temp_root = dir.path().join("temp-root");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                entries: vec![ParagraphEntry {
                    text: "Published through workflow".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "cyan".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let report = publish_spec_file_to_docx(&input_path, &output_path, &spec, dir.path(), Some(&temp_root))
            .expect("publish workflow should succeed");

        assert_eq!(report.published_output, output_path);
        assert!(report.xml_parts_checked > 0);
        assert!(output_path.exists());
        assert!(temp_root.exists());
        assert_eq!(
            fs::read_dir(&temp_root).expect("temp root").count(),
            0,
            "success path should clean temporary workspace"
        );

        let bytes = fs::read(&output_path).expect("read output");
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip");
        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document")
            .read_to_string(&mut document)
            .expect("read document");
        assert!(document.contains("Published through workflow"));
    }

    #[test]
    fn publish_workflow_fails_without_publishing_output() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("input.docx");
        let output_path = dir.path().join("published.docx");
        let temp_root = dir.path().join("temp-root");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("missing anchor".to_string()),
                entries: vec![ParagraphEntry {
                    text: "Should not publish".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "green".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let err = publish_spec_file_to_docx(&input_path, &output_path, &spec, dir.path(), Some(&temp_root))
            .expect_err("publish workflow should fail");

        assert!(
            err.to_string()
                .contains("publish workflow aborted; temporary workspace preserved"),
            "{err}"
        );
        assert!(!output_path.exists(), "failed workflow must not publish output");
        assert!(
            fs::read_dir(&temp_root).expect("temp root").count() >= 1,
            "failure path should preserve temp workspace for debugging"
        );
    }

    #[test]
    fn fidelity_check_detects_lost_highlight_in_untouched_paragraph() {
        let source = make_highlighted_docx().expect("source docx");
        let candidate = replace_document_xml(
            &source,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1" /></w:pPr>
      <w:r><w:t>Preserve me</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
    <w:sectPr />
  </w:body>
</w:document>"#,
        )
        .expect("candidate docx");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Anchor".to_string()),
                entries: vec![ParagraphEntry {
                    text: "New paragraph".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "cyan".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let report =
            validate_source_fidelity_bytes(&source, &candidate, &spec).expect("fidelity report");
        assert!(!report.is_valid());
        assert!(
            report
                .issues
                .iter()
                .any(|issue| issue.contains("Preserve me")),
            "{:?}",
            report.issues
        );
    }

    #[test]
    fn fidelity_check_ignores_replaced_paragraphs() {
        let source = make_highlighted_docx().expect("source docx");
        let candidate = replace_document_xml(
            &source,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1" /></w:pPr>
      <w:r><w:t>Updated text</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
    <w:sectPr />
  </w:body>
</w:document>"#,
        )
        .expect("candidate docx");

        let spec = AutomationSpec {
            operations: vec![Operation::ReplaceText {
                part: PartTarget::default(),
                find: "Preserve me".to_string(),
                replace: "Updated text".to_string(),
                highlight: "cyan".to_string(),
            }],
        };

        let report =
            validate_source_fidelity_bytes(&source, &candidate, &spec).expect("fidelity report");
        assert!(report.is_valid(), "{:?}", report.issues);
    }

    #[test]
    fn word_fidelity_exports_detect_lost_highlights() {
        let dir = tempdir().expect("temp dir");
        let source_path = dir.path().join("source.json");
        let migrated_path = dir.path().join("migrated.json");

        fs::write(
            &source_path,
            r#"[
  {"text":"Preserve me","style":"Heading 1","highlights":["7"]},
  {"text":"Anchor","style":null,"highlights":[]}
]"#,
        )
        .expect("write source");
        fs::write(
            &migrated_path,
            r#"[
  {"text":"Preserve me","style":"Heading 1","highlights":[]},
  {"text":"Anchor","style":null,"highlights":[]}
]"#,
        )
        .expect("write migrated");

        let report =
            validate_word_fidelity_exports(&source_path, &migrated_path).expect("report");
        assert!(!report.is_valid());
        assert!(
            report.issues.iter().any(|issue| issue.contains("Preserve me")),
            "{:?}",
            report.issues
        );
    }

    #[test]
    fn word_fidelity_exports_accept_matching_signatures() {
        let dir = tempdir().expect("temp dir");
        let source_path = dir.path().join("source.json");
        let migrated_path = dir.path().join("migrated.json");

        let content = r#"[
  {"text":"Preserve me","style":"Heading 1","highlights":["7"]},
  {"text":"Anchor","style":null,"highlights":[]}
]"#;
        fs::write(&source_path, content).expect("write source");
        fs::write(&migrated_path, content).expect("write migrated");

        let report =
            validate_word_fidelity_exports(&source_path, &migrated_path).expect("report");
        assert!(report.is_valid(), "{:?}", report.issues);
    }

    #[test]
    fn normalization_inspection_accepts_ooxml_zip() {
        let bytes = make_minimal_docx().expect("docx");
        let report = inspect_normalization_bytes(&bytes);
        assert_eq!(report.format, DocumentFormat::OoxmlZip);
        assert!(report.is_normalized);
        assert!(!report.requires_normalization);
    }

    #[test]
    fn normalization_inspection_detects_ole_encrypted_package() {
        let mut bytes = OLE_HEADER.to_vec();
        bytes.extend_from_slice(&utf16le_bytes("EncryptedPackage"));
        let report = inspect_normalization_bytes(&bytes);
        assert_eq!(report.format, DocumentFormat::OleEncryptedPackage);
        assert!(!report.is_normalized);
        assert!(report.requires_normalization);
    }

    #[test]
    fn normalize_workflow_passes_through_valid_ooxml() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("input.docx");
        let output_path = dir.path().join("normalized.docx");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");

        let report =
            normalize_docx_file(&input_path, &output_path, None, None, false).expect("normalize");
        assert_eq!(report.detected_format, DocumentFormat::OoxmlZip);
        assert!(report.already_normalized);
        assert!(report.xml_parts_checked > 0);
        assert_eq!(fs::read(&input_path).expect("read input"), fs::read(&output_path).expect("read output"));
    }

    #[test]
    fn verify_published_output_matches_candidate_accepts_identical_bytes() {
        let dir = tempdir().expect("temp dir");
        let output_path = dir.path().join("output.docx");
        let bytes = make_minimal_docx().expect("docx");
        fs::write(&output_path, &bytes).expect("write output");

        verify_published_output_matches_candidate(&output_path, &bytes)
            .expect("published output should match candidate");
    }

    #[test]
    fn verify_published_output_matches_candidate_rejects_mutated_output() {
        let dir = tempdir().expect("temp dir");
        let output_path = dir.path().join("output.docx");
        let candidate = make_minimal_docx().expect("docx");
        let mut mutated = OLE_HEADER.to_vec();
        mutated.extend_from_slice(&utf16le_bytes("EncryptedPackage"));
        fs::write(&output_path, &mutated).expect("write mutated output");

        let err = verify_published_output_matches_candidate(&output_path, &candidate)
            .expect_err("published output should be rejected");
        let message = format!("{err:#}");
        assert!(message.contains("published output mutated after write"));
        assert!(message.contains("ole-encrypted-package"));
    }

    #[test]
    fn spec_validation_rejects_unsupported_highlight() {
        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                entries: vec![ParagraphEntry {
                    text: "Invalid color".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "pink".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let err = spec.validate(Path::new(".")).expect_err("spec should fail validation");
        assert!(err.to_string().contains("unsupported highlight 'pink'"), "{err}");
    }

    #[test]
    fn prepare_work_session_reuses_cached_normalized_copy() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("strategy-v007.docx");
        let temp_root = dir.path().join("temp-root");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");

        let first =
            prepare_work_session(&input_path, "strategy", Some(&temp_root), None, false).expect("first session");
        let second =
            prepare_work_session(&input_path, "strategy", Some(&temp_root), None, false).expect("second session");

        assert!(!first.cache_hit);
        assert!(first.normalized_input.exists());
        assert!(second.cache_hit);
        assert_eq!(first.normalized_input, second.normalized_input);
    }

    #[test]
    fn publish_next_uses_highest_existing_version() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("versioned-document-v007.docx");
        let existing_path = dir.path().join("versioned-document-v010.docx");
        let temp_root = dir.path().join("temp-root");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");
        fs::write(&existing_path, make_minimal_docx().expect("docx")).expect("write existing");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                entries: vec![ParagraphEntry {
                    text: "Next version paragraph".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "magenta".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let report = publish_spec_file_to_next_version(
            &input_path,
            &spec,
            dir.path(),
            Some(&temp_root),
            None,
            PublishTargetMode::NextVersion,
        )
        .expect("publish next");

        assert_eq!(report.version_number, 11);
        assert!(report.published_output.ends_with("versioned-document-v011.docx"));
        assert!(report.published_output.exists());
        assert_eq!(report.mode, PublishTargetMode::NextVersion);
    }

    #[test]
    fn publish_next_latest_updates_current_version_in_place() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("versioned-document-v007.docx");
        let existing_path = dir.path().join("versioned-document-v010.docx");
        let temp_root = dir.path().join("temp-root");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");
        fs::write(&existing_path, make_minimal_docx().expect("docx")).expect("write existing");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                entries: vec![ParagraphEntry {
                    text: "Continue working paragraph".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "green".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let report = publish_spec_file_to_next_version(
            &input_path,
            &spec,
            dir.path(),
            Some(&temp_root),
            None,
            PublishTargetMode::Latest,
        )
        .expect("publish latest");

        assert_eq!(report.version_number, 10);
        assert!(report.published_output.ends_with("versioned-document-v010.docx"));
        assert_eq!(report.mode, PublishTargetMode::Latest);
    }

    #[test]
    fn publish_next_latest_rejects_existing_non_normalized_output() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("versioned-document-v007.docx");
        let existing_path = dir.path().join("versioned-document-v010.docx");
        let temp_root = dir.path().join("temp-root");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");
        let mut protected = OLE_HEADER.to_vec();
        protected.extend_from_slice(&utf16le_bytes("EncryptedPackage"));
        fs::write(&existing_path, protected).expect("write protected existing");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                entries: vec![ParagraphEntry {
                    text: "Continue working paragraph".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "green".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let err = publish_spec_file_to_next_version(
            &input_path,
            &spec,
            dir.path(),
            Some(&temp_root),
            None,
            PublishTargetMode::Latest,
        )
        .expect_err("latest publish should reject protected output");
        let message = format!("{err:#}");
        assert!(message.contains("refusing to overwrite existing non-normalized"));
        assert!(message.contains("ole-encrypted-package"));
    }

    #[test]
    fn publish_session_to_next_version_updates_session_progress() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("versioned-document-v007.docx");
        let temp_root = dir.path().join("temp-root");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");

        prepare_work_session(&input_path, "versioned-document", Some(&temp_root), None, false)
            .expect("prepare session");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                entries: vec![ParagraphEntry {
                    text: "Session paragraph".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "green".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let first = publish_session_to_next_version(
            "versioned-document",
            &spec,
            dir.path(),
            Some(&temp_root),
            None,
            PublishTargetMode::NextVersion,
        )
        .expect("session publish");
        let metadata = load_work_session_metadata(&temp_root, "versioned-document").expect("metadata");

        assert_eq!(first.version_number, 8);
        assert!(metadata.current_version_path.ends_with("versioned-document-v008.docx"));
        assert!(Path::new(&metadata.normalized_path).exists());
        assert_eq!(first.mode, PublishTargetMode::NextVersion);
    }

    #[test]
    fn publish_session_latest_keeps_same_version_path() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("versioned-document-v007.docx");
        let temp_root = dir.path().join("temp-root");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");

        prepare_work_session(&input_path, "versioned-document-latest", Some(&temp_root), None, false)
            .expect("prepare session");

        let spec = AutomationSpec {
            operations: vec![Operation::InsertParagraphs {
                part: PartTarget::default(),
                anchor: AnchorTarget::Simple("Strategic principles".to_string()),
                entries: vec![ParagraphEntry {
                    text: "Latest version paragraph".to_string(),
                    style: ParagraphStyle::Normal,
                    highlight: "green".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                }],
            }],
        };

        let report = publish_session_to_next_version(
            "versioned-document-latest",
            &spec,
            dir.path(),
            Some(&temp_root),
            None,
            PublishTargetMode::Latest,
        )
        .expect("session latest");
        let metadata =
            load_work_session_metadata(&temp_root, "versioned-document-latest").expect("metadata");

        assert_eq!(report.version_number, 7);
        assert!(report.published_output.ends_with("versioned-document-v007.docx"));
        assert!(metadata.current_version_path.ends_with("versioned-document-v007.docx"));
        assert_eq!(report.mode, PublishTargetMode::Latest);
    }

    #[test]
    fn comment_crud_supports_list_update_and_delete() {
        let dir = tempdir().expect("temp dir");
        let input_path = dir.path().join("input.docx");
        let commented_path = dir.path().join("commented.docx");
        let updated_path = dir.path().join("updated.docx");
        let deleted_path = dir.path().join("deleted.docx");
        fs::write(&input_path, make_minimal_docx().expect("docx")).expect("write input");

        let comment = CommentSpec {
            text: "Copilot review".to_string(),
            comment_text: "Please clarify the approval workflow.".to_string(),
            author: "GitHub Copilot CLI".to_string(),
            initials: Some("GCC".to_string()),
            date: Some("2026-05-20T00:00:00Z".to_string()),
            style: ParagraphStyle::Normal,
            highlight: "yellow".to_string(),
        };
        add_comment_to_docx(
            &input_path,
            &commented_path,
            None,
            &AnchorTarget::Simple("Strategic principles".to_string()),
            &comment,
        )
        .expect("add comment");

        let listed = list_docx_comments(&commented_path).expect("list comments");
        assert_eq!(listed.len(), 1);
        assert_eq!(listed[0].author, "GitHub Copilot CLI");
        assert_eq!(listed[0].comment_text, "Please clarify the approval workflow.");
        assert!(
            listed[0]
                .locations
                .iter()
                .any(|location| location.paragraph_text.contains("Strategic principles"))
        );

        update_docx_comment(
            &commented_path,
            &updated_path,
            listed[0].id,
            &CommentUpdate {
                comment_text: Some("Please clarify the review workflow.".to_string()),
                author: Some("GitHub Copilot CLI".to_string()),
                initials: None,
                date: None,
                highlight: Some("magenta".to_string()),
            },
        )
        .expect("update comment");
        let updated = list_docx_comments(&updated_path).expect("list updated comments");
        assert_eq!(updated[0].comment_text, "Please clarify the review workflow.");
        assert_eq!(updated[0].highlight, "magenta");

        delete_docx_comment(&updated_path, &deleted_path, updated[0].id).expect("delete comment");
        let deleted = list_docx_comments(&deleted_path).expect("list deleted comments");
        assert!(deleted.is_empty());
        let bytes = fs::read(&deleted_path).expect("read deleted output");
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip");
        let mut document = String::new();
        archive
            .by_name(DOCUMENT_XML_PATH)
            .expect("document xml")
            .read_to_string(&mut document)
            .expect("read document");
        assert!(!document.contains("commentRangeStart"));
        assert!(!document.contains("commentReference"));
    }

    fn make_minimal_docx() -> Result<Vec<u8>> {
        let mut out = Cursor::new(Vec::new());
        {
            let mut writer = ZipWriter::new(&mut out);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

            writer.start_file(CONTENT_TYPES_XML_PATH, options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
            )?;

            writer.add_directory("_rels/", options)?;
            writer.start_file(ROOT_RELS_XML_PATH, options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
            )?;

            writer.add_directory("word/", options)?;
            writer.start_file(DOCUMENT_XML_PATH, options)?;
            writer.write_all(MINIMAL_DOCUMENT_XML.as_bytes())?;
            writer.finish()?;
        }

        Ok(out.into_inner())
    }

    fn make_highlighted_docx() -> Result<Vec<u8>> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1" /></w:pPr>
      <w:r><w:rPr><w:highlight w:val="yellow" /></w:rPr><w:t>Preserve me</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
    <w:sectPr />
  </w:body>
</w:document>"#;
        make_docx_with_document_xml(xml)
    }

    fn make_docx_with_document_xml(document_xml: &[u8]) -> Result<Vec<u8>> {
        let mut out = Cursor::new(Vec::new());
        {
            let mut writer = ZipWriter::new(&mut out);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

            writer.start_file(CONTENT_TYPES_XML_PATH, options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
            )?;

            writer.add_directory("_rels/", options)?;
            writer.start_file(ROOT_RELS_XML_PATH, options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
            )?;

            writer.add_directory("word/", options)?;
            writer.start_file(DOCUMENT_XML_PATH, options)?;
            writer.write_all(document_xml)?;
            writer.finish()?;
        }
        Ok(out.into_inner())
    }

    fn replace_document_xml(source: &[u8], document_xml: &str) -> Result<Vec<u8>> {
        let mut package = DocxPackage::read(source)?;
        package.set_file(DOCUMENT_XML_PATH, document_xml.as_bytes().to_vec());
        package.write()
    }

    fn make_repeated_anchor_docx() -> Result<Vec<u8>> {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Strategic principles</w:t></w:r></w:p>
    <w:p><w:r><w:t>Other paragraph</w:t></w:r></w:p>
    <w:p><w:r><w:t>Strategic principles</w:t></w:r></w:p>
    <w:sectPr />
  </w:body>
</w:document>"#;

        let mut out = Cursor::new(Vec::new());
        {
            let mut writer = ZipWriter::new(&mut out);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            writer.start_file(CONTENT_TYPES_XML_PATH, options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
            )?;
            writer.add_directory("_rels/", options)?;
            writer.start_file(ROOT_RELS_XML_PATH, options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
            )?;
            writer.add_directory("word/", options)?;
            writer.start_file(DOCUMENT_XML_PATH, options)?;
            writer.write_all(xml)?;
            writer.finish()?;
        }
        Ok(out.into_inner())
    }

    fn make_word_like_docx() -> Result<Vec<u8>> {
        let document_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" mc:Ignorable="w14 wp14">
  <w:body>
    <w:p><w:r><w:t>Animated GIF sample document</w:t></w:r></w:p>
    <w:p><w:r><w:t>Strategic principles</w:t></w:r></w:p>
    <w:sectPr />
  </w:body>
</w:document>"#;

        let mut out = Cursor::new(Vec::new());
        {
            let mut writer = ZipWriter::new(&mut out);
            let options =
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            writer.start_file(CONTENT_TYPES_XML_PATH, options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#,
            )?;
            writer.add_directory("_rels/", options)?;
            writer.start_file(ROOT_RELS_XML_PATH, options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
            )?;
            writer.add_directory("word/", options)?;
            writer.start_file(DOCUMENT_XML_PATH, options)?;
            writer.write_all(document_xml)?;
            writer.start_file("word/styles.xml", options)?;
            writer.write_all(
                br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>"#,
            )?;
            writer.finish()?;
        }
        Ok(out.into_inner())
    }

    fn utf16le_bytes(value: &str) -> Vec<u8> {
        let mut bytes = Vec::new();
        for unit in value.encode_utf16() {
            bytes.extend_from_slice(&unit.to_le_bytes());
        }
        bytes
    }
}
