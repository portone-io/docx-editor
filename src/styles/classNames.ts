/**
 * The CSS class and variable names attached to the UI.
 *
 * They are built only here, so that renaming the package means changing a single
 * prefix line.
 * `editor.css` uses the same prefix, so change it along with this file.
 */

const PREFIX = "docx-editor";

export const editorClassNames = {
  /** The outermost box holding both the toolbar and the paper */
  frame: `${PREFIX}-frame`,
  /** The box that owns the scrolling */
  root: `${PREFIX}-editor`,
  /** The paper the text sits on (contenteditable) */
  sheet: `${PREFIX}-sheet`,
  paragraph: `${PREFIX}-p`,
  run: `${PREFIX}-run`,
  tab: `${PREFIX}-tab`,
  /** The independently measurable DOM range around one tab character. */
  tabSlot: `${PREFIX}-tab-slot`,
  /** The caret drawn at an unambiguous edge of a laid-out tab. */
  tabCaret: `${PREFIX}-tab-caret`,
  /** Attached to the contenteditable while its native caret is replaced. */
  tabCaretActive: `${PREFIX}-tab-caret-active`,
  /** An inline image (`w:drawing` holding a picture) */
  image: `${PREFIX}-img`,
  /** The box the image sits in, which is also what the resize handles stand in */
  imageBox: `${PREFIX}-img-box`,
  /** Attached to that box while the image is the selection */
  imageSelected: `${PREFIX}-img-selected`,
  /** One of the four corner handles a selected image is resized by */
  imageHandle: `${PREFIX}-img-handle`,
  table: `${PREFIX}-tbl`,
  tableRow: `${PREFIX}-tr`,
  tableCell: `${PREFIX}-tc`,
  /** Attached alongside it when the control around the cell says the cell may not be edited */
  cellLocked: `${PREFIX}-tc-locked`,
  /** A stretch of text sitting inside a content control (`w:sdt`) */
  sdt: `${PREFIX}-sdt`,
  /** Attached alongside it when the control says its contents may not be edited */
  sdtLocked: `${PREFIX}-sdt-locked`,
  /** A stretch of text inside a hyperlink (`w:hyperlink`) */
  link: `${PREFIX}-link`,
  /** A stretch of text carrying a comment anchor */
  commentRange: `${PREFIX}-comment-range`,
  /** An invisible OOXML comment marker inside the editable text flow */
  commentMarker: `${PREFIX}-comment-marker`,
  /** A superscript footnote or endnote reference */
  noteReference: `${PREFIX}-note-reference`,
  /** An invisible bookmark marker that occurs between body blocks */
  bookmarkBlock: `${PREFIX}-bookmark-block`,
  rawInline: `${PREFIX}-raw-inline`,
  rawBlock: `${PREFIX}-raw-block`,
  /** A preserved block that carries the original XML directly (inside a table cell, and so on) */
  rawXmlBlock: `${PREFIX}-raw-xml`,
  tablePlaceholder: `${PREFIX}-table`,
  /** The box holding both the paper and the marks overlaid on it */
  pageLayer: `${PREFIX}-page-layer`,
  /** All of the page marks drawn over the paper */
  pageGuides: `${PREFIX}-page-guides`,
  /** The band that covers the space between one page and the next */
  pageSplit: `${PREFIX}-page-split`,
  /** The line drawn where a block that cannot be pushed down crosses a page */
  pageCrossed: `${PREFIX}-page-crossed`,
  /** The page number laid on the top corner of a page */
  pageBadge: `${PREFIX}-page-badge`,
  /** A first-section header story projected into a visual page margin */
  pageHeader: `${PREFIX}-page-header`,
  /** A first-section footer story projected into a visual page margin */
  pageFooter: `${PREFIX}-page-footer`,
  /** The mark that changes the cursor while it is over a column boundary or while a boundary is being dragged. Attached to the sheet */
  columnResize: `${PREFIX}-col-resize`,
  /** The vertical line that shows where the boundary stands while it is being dragged */
  columnResizeGuide: `${PREFIX}-col-guide`,
  rowResize: `${PREFIX}-row-resize`,
  rowResizeGuide: `${PREFIX}-row-guide`,

  /* The controls that come with the editor */
  toolbar: `${PREFIX}-toolbar`,
  /** A bundle of buttons set off by a single separator */
  toolbarGroup: `${PREFIX}-toolbar-group`,
  toolbarSeparator: `${PREFIX}-toolbar-sep`,
  toolbarButton: `${PREFIX}-toolbar-btn`,
  toolbarColorIcon: `${PREFIX}-toolbar-color-icon`,
  toolbarColorIndicator: `${PREFIX}-toolbar-color-indicator`,
  toolbarSelect: `${PREFIX}-toolbar-select`,
  fontFamilySelect: `${PREFIX}-font-family-select`,
  zoomSelect: `${PREFIX}-zoom-select`,
  /** The box wrapped around a control that cannot draw its own tooltip (a native `select`) */
  tooltipWrapper: `${PREFIX}-tooltip-wrap`,
  /** The box around the button a panel opens from. A click inside it counts as a click on the panel */
  popoverBox: `${PREFIX}-popover-box`,
  /** The small panel that opens below a button, floated against the screen */
  popover: `${PREFIX}-popover`,
  /** The grid for picking a table size */
  sizeGrid: `${PREFIX}-size-grid`,
  sizeGridRow: `${PREFIX}-size-row`,
  sizeGridCell: `${PREFIX}-size-cell`,
  sizeGridLabel: `${PREFIX}-size-label`,
  colorPicker: `${PREFIX}-color-picker`,
  /** The text color palette. A grid of rows, so the markup says what the layout shows */
  colorGrid: `${PREFIX}-color-grid`,
  /** One row of the palette. Ten swatches of a family, light to dark */
  colorRow: `${PREFIX}-color-row`,
  /** The row at the head of the palette holding the entry that withdraws the color */
  colorClearRow: `${PREFIX}-color-clear-row`,
  swatch: `${PREFIX}-swatch`,
  colorCustom: `${PREFIX}-color-custom`,
  colorNativeInput: `${PREFIX}-color-native-input`,
  colorHexInput: `${PREFIX}-color-hex-input`,
  colorApply: `${PREFIX}-color-apply`,
  /** The table menu opened with the right button, and the choice menus inside the toolbar panels */
  menu: `${PREFIX}-menu`,
  /** A menu of choices inside a toolbar panel, where one of them is the current one */
  menuList: `${PREFIX}-menu-list`,
  menuItem: `${PREFIX}-menu-item`,
  /** The slot at the head of a menu row where the check mark of the current choice goes */
  menuCheck: `${PREFIX}-menu-check`,
  /** The keyboard shortcut written on the right of a menu row */
  menuHint: `${PREFIX}-menu-hint`,
  menuSeparator: `${PREFIX}-menu-sep`,
  /** The panel that reads and writes a link's address, floated over the selection */
  linkPanel: `${PREFIX}-link-panel`,
  /** The field the address is typed into */
  linkAddress: `${PREFIX}-link-address`,
  /** The row of buttons under it */
  linkPanelRow: `${PREFIX}-link-row`,
  linkPanelIcon: `${PREFIX}-link-icon`,
  /** A button of that panel, and of the card below, which offer the same actions */
  /** The card that says where the link under the cursor points, floated under the link's first line */
  linkCard: `${PREFIX}-link-card`,
  /** The address it shows, cut off with an ellipsis where it is longer than the card */
  linkCardAddress: `${PREFIX}-link-card-address`,
  /** What it says in place of an address for a link that names a place in the document */
  linkCardNote: `${PREFIX}-link-card-note`,
  /** The document and its comment rail, laid out side by side */
  workspace: `${PREFIX}-workspace`,
  commentsPanel: `${PREFIX}-comments`,
  commentsToggle: `${PREFIX}-comments-toggle`,
  commentsHeading: `${PREFIX}-comments-heading`,
  commentsCanvas: `${PREFIX}-comments-canvas`,
  commentsEmpty: `${PREFIX}-comments-empty`,
  commentPosition: `${PREFIX}-comment-position`,
  commentCard: `${PREFIX}-comment-card`,
  commentHeader: `${PREFIX}-comment-header`,
  commentReply: `${PREFIX}-comment-reply`,
  commentMeta: `${PREFIX}-comment-meta`,
  commentAuthor: `${PREFIX}-comment-author`,
  commentDate: `${PREFIX}-comment-date`,
  commentBody: `${PREFIX}-comment-body`,
  commentActions: `${PREFIX}-comment-actions`,
  commentPrimaryAction: `${PREFIX}-comment-primary-action`,
  commentIconActions: `${PREFIX}-comment-icon-actions`,
  commentComposer: `${PREFIX}-comment-composer`,
  commentInput: `${PREFIX}-comment-input`,
  notesPanel: `${PREFIX}-notes`,
  notesHeading: `${PREFIX}-notes-heading`,
  notesList: `${PREFIX}-notes-list`,
  noteLabel: `${PREFIX}-note-label`,
  noteBody: `${PREFIX}-note-body`,
  /** The panel drawn in place of the editor when a document was refused and the consumer wrote no panel of its own */
  rejection: `${PREFIX}-rejection`,
  rejectionTitle: `${PREFIX}-rejection-title`,
} as const;

export const editorCssVariables = {
  fontSize: `--${PREFIX}-font-size`,
  fontFamily: `--${PREFIX}-font-family`,
  lineHeight: `--${PREFIX}-line-height`,
  /** The hanging indent width the list marker sits in */
  markerWidth: `--${PREFIX}-marker-width`,
  /** The width calculated for a tab using custom paragraph stops. */
  tabWidth: `--${PREFIX}-tab-width`,
  /** The paper height, grown to match the page count */
  sheetHeight: `--${PREFIX}-sheet-height`,
  /** The paper the open document names, and the margins printed inside it */
  pageWidth: `--${PREFIX}-page-width`,
  pageHeight: `--${PREFIX}-page-height`,
  pageMarginTop: `--${PREFIX}-page-margin-top`,
  pageMarginRight: `--${PREFIX}-page-margin-right`,
  pageMarginBottom: `--${PREFIX}-page-margin-bottom`,
  pageMarginLeft: `--${PREFIX}-page-margin-left`,
} as const;

/**
 * The attributes that carry marks which exist only on screen.
 * A list marker is a decoration rather than part of the document model, so CSS
 * draws it from this attribute alone.
 * The boundary guides use these attributes to find the places that ask for a page
 * break.
 */
export const editorAttributes = {
  listMarker: "data-marker",
  /** A new page starts at this paragraph */
  pageBreakBefore: "data-page-break",
  /** The kind of line break. `page` when the break splits the page */
  breakType: "data-break",
  /** Which corner of the image a resize handle stands at (`nw`, `ne`, `sw`, `se`) */
  imageCorner: "data-corner",
  /** The extra height opened up to push content down to the top of the next page */
  pagePush: "data-page-push",
  /** The space opened up at a page break, so what follows the break starts the next page */
  pageBreakSpace: "data-page-space",
  /** A display-only table row that fills the remainder between page bodies */
  tablePageSpace: "data-table-page-space",
  /** A display-only copy of a table header row on a continued page */
  tableRepeatedHeader: "data-table-repeated-header",
  /**
   * The tooltip a control shows on hover, drawn by the CSS straight from this attribute.
   * It is the one mark here whose selector is the attribute on its own rather than a
   * class of ours, so it carries the prefix and cannot collide with the host app.
   */
  tooltip: `data-${PREFIX}-tooltip`,
  /**
   * Put on the box the tooltips are drawn inside while they are hushed, which is what Escape
   * does to the one standing in the way (`ui/Tooltip.tsx`).
   */
  tooltipsHushed: `data-${PREFIX}-tooltips-off`,
} as const;
