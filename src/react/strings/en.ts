/**
 * The viewer's UI texts in English (the default locale). Every other locale has the same keys (`Strings`);
 * values are plain strings, or functions where a text takes numbers or names.
 */
export const en = {
  /** BCP 47 tag used for numbers and dates (Intl) */
  locale: 'en-US',

  common: {
    ok: 'OK',
    cancel: 'Cancel',
    close: 'Close',
    save: 'Save',
    delete: 'Delete',
    edit: 'Edit',
    more: 'More',
    none: 'None',
    apply: 'Apply',
    remove: 'Remove',
    insert: 'Insert',
    loading: 'Loading…',
    stillLoading: 'The document is still loading — try again in a moment',
  },

  bar: {
    cannotOpen: 'Cannot open document.',
    downloading: (loaded: string, total: string) => `Downloading… ${loaded} / ${total}`,
    opening: 'Opening…',
    noDocument: 'No document',
    code: 'Code',
    codeHide: 'Hide the source',
    codeShow: 'Edit the source side by side with the document',
    comments: 'Comments',
    commentsTitle: (keys: string) => `Comments (${keys} to comment on the selection)`,
    view: 'View',
    edit: 'Edit',
    toView: 'Switch to viewing',
    toEdit: 'Edit document',
    saveTitle: 'Save a copy (Ctrl+S)',
    saving: 'Saving…',
    save: 'Save',
    download: 'Download',
    sourceStale: 'Changes made in the document are not in this source: editing it builds the document again from the source.',
    sourceCopied: 'Source copied',
    cannotCopy: 'Cannot copy here',
    resizeSource: 'Resize the source',
    discardTitle: 'Discard unsaved changes?',
    discardOk: 'Continue',
    replaceNew: (what: string) => `A new ${what} replaces the current document. Your changes are not saved.`,
    replaceOpen: (name: string) => `Opening “${name}” replaces the current document. Your changes are not saved.`,
    renameTitle: 'Rename',
    renameLabel: 'File name',
    renameOk: 'Rename',
  },

  /** file names of new documents (the extension is added) */
  untitled: { docx: 'Untitled document', xlsx: 'Untitled spreadsheet', pptx: 'Untitled presentation' },
  /** what File → New makes, in "A new … replaces the current document" */
  newWhat: { docx: 'document', xlsx: 'spreadsheet', pptx: 'presentation' },

  kind: { docx: 'Word document', xlsx: 'Excel workbook', pptx: 'PowerPoint presentation' },
  from: { pages: 'Apple Pages document', keynote: 'Apple Keynote presentation', numbers: 'Apple Numbers spreadsheet' },
  converted: {
    doc: 'Word 97-2003 document',
    xls: 'Excel 97-2003 workbook',
    ppt: 'PowerPoint 97-2003 presentation',
    rtf: 'Rich Text Format document',
    odt: 'OpenDocument text',
    ods: 'OpenDocument spreadsheet',
    odp: 'OpenDocument presentation',
  },
  imported: {
    csv: 'CSV file',
    tsv: 'Tab-separated values',
    text: 'Plain text',
    markdown: 'Markdown',
    html: 'HTML page',
    svg: 'SVG picture',
    mermaid: 'Mermaid diagram',
    plantuml: 'PlantUML diagram',
    graphviz: 'Graphviz diagram',
  },

  details: {
    title: 'Document details',
    name: 'Name',
    type: 'Type',
    viewOnly: (what: string) => `${what} (view only)`,
    openedAs: (what: string, kind: string) => `${what}, opened as ${kind}`,
    fileSize: 'File size',
    status: 'Status',
    unsaved: 'Unsaved changes',
    core: {
      title: 'Title',
      subject: 'Subject',
      creator: 'Author',
      keywords: 'Keywords',
      description: 'Comments',
      lastModifiedBy: 'Last modified by',
      created: 'Created',
      modified: 'Modified',
    },
  },

  file: {
    menu: 'File',
    new: 'New',
    newDocx: 'Document',
    newXlsx: 'Spreadsheet',
    newPptx: 'Presentation',
    open: 'Open…',
    importCsv: 'Import CSV into new sheet…',
    save: 'Save',
    rename: 'Rename…',
    downloadAs: 'Download as',
    pdf: 'PDF document',
    html: 'Web page',
    txt: 'Plain text',
    md: 'Markdown',
    csv: 'CSV (current sheet)',
    tsv: 'TSV (current sheet)',
    picture: 'SVG picture',
    details: 'Document details',
    print: 'Print',
    close: 'Close',
    preparingPdf: 'Preparing PDF… choose “Save as PDF” in the print dialog',
    preparing: 'Preparing…',
    unsupported: (name: string) =>
      `Cannot open “${name}”: supported files are Word, Excel, PowerPoint (also 97-2003), OpenDocument, RTF, Pages, Keynote, Numbers, CSV, text, Markdown, HTML, SVG, Mermaid, PlantUML and Graphviz (also gzip-compressed: .svgz, .csv.gz)`,
  },

  zoom: {
    group: 'Zoom',
    out: 'Zoom out',
    in: 'Zoom in',
    fit: 'Fit',
    fitTitle: 'Fit width (click for 100%)',
    toFit: 'Click to fit width',
  },

  toolbar: {
    fontSize: 'Font size',
    fontSizeDown: 'Decrease font size',
    fontSizeUp: 'Increase font size',
    more: 'More',
    moreTools: 'More tools',
    custom: 'Custom',
    customColor: 'Custom color',
    customColorTitle: 'Custom color…',
    tableSize: 'Table size',
  },

  source: {
    source: 'Source',
    lines: (n: number) => (n === 1 ? 'line' : 'lines'),
    copy: 'Copy the source',
    download: 'Download the source',
    input: (format: string) => `${format} source`,
  },

  picture: {
    notShown: 'The picture is not shown',
    size: 'Picture size',
    notDrawn: 'SVG picture not drawn.',
  },

  collab: {
    connecting: 'Connecting…',
    syncing: 'Syncing…',
    live: 'Live',
    offline: 'Offline · changes will sync',
    error: 'Not synced',
    stopped: 'Collaboration stopped',
    saving: (n: number) => `Saving ${n} change${n > 1 ? 's' : ''}…`,
    editing: (n: number) => `${n} editing`,
  },
}

export type Strings = typeof en
