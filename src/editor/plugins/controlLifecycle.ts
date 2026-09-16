/**
 * What a content control says about outliving an edit of its contents, appended to the edit that
 * caused it (`spec/notes/contentControls.md`).
 *
 * Neither `w:temporary` nor `w:showingPlcHdr` can be answered by preserving the control's opening
 * XML, so the first edit inside a control carrying one lifts the wrapper or rewrites that XML
 * without the flag. Appending it is what puts it in the history beside the edit, so one undo takes
 * the two back together.
 */

import { type Mark, type Node as PMNode, Slice } from "prosemirror-model";
import { Plugin, type Transaction } from "prosemirror-state";
import { Mapping, ReplaceAroundStep } from "prosemirror-transform";
import { editSdtPrefix } from "../../docx/sdt";
import {
  type ControlAttrNames,
  type ControlFacts,
  controlAttrs,
  controlAttrsOf,
  controlFactsOf,
  isBlockControl,
  NO_CONTROL,
  OWN_CONTROL_ATTRS,
} from "../../schema/controlAttrs";
import { displayOnly } from "../../schema/displayDerivation";
import { controlLifted, controlSpans, historyReplay } from "../../schema/locks";
import { changesOnlyComments } from "../../schema/protection";

interface Range {
  from: number;
  to: number;
}

/**
 * The stretches of the new document these transactions rewrote.
 *
 * Each step's own map names them in the document that step made, so what follows it is mapped over
 * them to bring every stretch into the coordinates of the document that came out at the end.
 */
function rewritten(transactions: readonly Transaction[]): Range[] {
  const whole = new Mapping();
  for (const transaction of transactions)
    whole.appendMapping(transaction.mapping);
  const ranges: Range[] = [];
  let index = 0;
  for (const transaction of transactions) {
    for (const map of transaction.mapping.maps) {
      const rest = whole.slice(index + 1);
      map.forEach((_from, _to, start, end) => {
        ranges.push({ from: rest.map(start, -1), to: rest.map(end, 1) });
      });
      index += 1;
    }
  }
  return ranges;
}

function edited(ranges: readonly Range[], from: number, to: number): boolean {
  return ranges.some((range) => range.from <= to && range.to >= from);
}

/** The opening XML with the placeholder flag taken out. null where it cannot be rewritten */
function withoutPlaceholderFlag(prefix: string | null): string | null {
  return prefix === null
    ? null
    : editSdtPrefix(prefix, [["showingPlcHdr", null]]);
}

/** Whether either property has anything to say about an edit inside this control */
function acts(control: ControlFacts): boolean {
  return control.temporary || control.showingPlaceholder;
}

/**
 * Whether the wrapper goes. A lock against deletion states that the control may not be removed
 * (§17.5.2.23) where `w:temporary` states that it must be, and the lock wins; it says nothing
 * about the placeholder flag, which is dropped either way.
 */
function lifts(control: ControlFacts): boolean {
  return control.temporary && !control.deletionLocked;
}

/**
 * Takes the wrapper away from one container, or rewrites its opening XML without the placeholder
 * flag. A block control is replaced by the blocks it holds; a wrapped cell or row keeps the node
 * and loses what the control said about it.
 */
function settleContainer(
  tr: Transaction,
  node: PMNode,
  pos: number,
  names: ControlAttrNames,
  control: ControlFacts
): void {
  const at = tr.mapping.map(pos);
  const end = tr.mapping.map(pos + node.nodeSize);
  if (lifts(control)) {
    if (isBlockControl(node)) {
      // The blocks stay put as the step's gap, so nothing locked inside is planted anywhere
      tr.step(
        new ReplaceAroundStep(at, end, at + 1, end - 1, Slice.empty, 0, true)
      );
      return;
    }
    tr.setNodeMarkup(at, undefined, {
      ...node.attrs,
      ...controlAttrs(names, NO_CONTROL),
    });
    return;
  }
  if (!control.showingPlaceholder) return;
  const prefix = withoutPlaceholderFlag(control.prefix);
  if (prefix === null) return;
  tr.setNodeMarkup(at, undefined, {
    ...node.attrs,
    ...controlAttrs(names, { ...control, prefix, showingPlaceholder: false }),
  });
}

/** The same for an inline control, which is a mark over the stretch of runs it wrapped */
function settleMark(
  tr: Transaction,
  span: { from: number; to: number; mark: Mark },
  control: ControlFacts
): void {
  const from = tr.mapping.map(span.from);
  const to = tr.mapping.map(span.to);
  if (lifts(control)) {
    tr.removeMark(from, to, span.mark);
    return;
  }
  if (!control.showingPlaceholder) return;
  const prefix = withoutPlaceholderFlag(control.prefix);
  if (prefix === null) return;
  tr.removeMark(from, to, span.mark);
  tr.addMark(
    from,
    to,
    span.mark.type.create({
      ...span.mark.attrs,
      ...controlAttrs(OWN_CONTROL_ATTRS, {
        ...control,
        prefix,
        showingPlaceholder: false,
      }),
    })
  );
}

function settleDocument(tr: Transaction, doc: PMNode, ranges: Range[]): void {
  doc.descendants((node, pos) => {
    // A carrier with no content has no interior an edit could reach - the two ends `edited` is
    // asked about would come out inverted - so neither property acts
    const names = node.content.size > 0 ? controlAttrsOf(node) : null;
    if (names) {
      const container = controlFactsOf(names, node.attrs);
      if (acts(container) && edited(ranges, pos + 1, pos + node.nodeSize - 1)) {
        settleContainer(tr, node, pos, names, container);
      }
    }
    if (!node.isTextblock) return true;
    for (const span of controlSpans({ node, start: pos + 1 })) {
      const inline = controlFactsOf(OWN_CONTROL_ATTRS, span.mark.attrs);
      if (!acts(inline)) continue;
      if (edited(ranges, span.from, span.to)) settleMark(tr, span, inline);
    }
    return false;
  });
}

/**
 * The passes a change may carry that say it is no edit of the contents.
 *
 * A history replay is the reverse of a step whose consequence was applied beside it and undone
 * beside it; a re-derivation of display values writes no content at all
 * (`schema/displayDerivation`); and what this plugin appends is the consequence rather than a
 * cause of one.
 */
const NOT_AN_EDIT = [historyReplay, displayOnly, controlLifted];

function anEdit(transaction: Transaction): boolean {
  return (
    transaction.docChanged &&
    !NOT_AN_EDIT.some((pass) => transaction.getMeta(pass) === true)
  );
}

/**
 * Applies both properties after every change that reached inside a control carrying one.
 *
 * The appended transaction carries the pass that lets a wrapper be lifted at all
 * (`schema/locks`): a control's edge and its lock both stand against the step that removes it,
 * and this is the one change neither speaks to.
 */
export function controlLifecycle(): Plugin {
  return new Plugin({
    appendTransaction(transactions, oldState, newState) {
      const changed = transactions.filter(anEdit);
      if (changed.length === 0) return null;
      const tr = newState.tr.setMeta(controlLifted, true);
      settleDocument(tr, newState.doc, rewritten(changed));
      if (!tr.docChanged) return null;
      return changesOnlyComments(oldState.doc, newState.doc) ? null : tr;
    },
  });
}
