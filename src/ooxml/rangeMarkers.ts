/** `EG_RangeMarkupElements` less the comment markers modelled inside paragraphs and wrappers */
export const RANGE_MARKERS = [
  "bookmarkStart",
  "bookmarkEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "customXmlInsRangeStart",
  "customXmlInsRangeEnd",
  "customXmlDelRangeStart",
  "customXmlDelRangeEnd",
  "customXmlMoveFromRangeStart",
  "customXmlMoveFromRangeEnd",
  "customXmlMoveToRangeStart",
  "customXmlMoveToRangeEnd",
];

/** The markers of `EG_RunLevelElts` that stand outside `EG_RangeMarkupElements` */
export const PERMISSION_MARKERS = ["permStart", "permEnd"];

export const COMMENT_RANGE_MARKERS = ["commentRangeStart", "commentRangeEnd"];
