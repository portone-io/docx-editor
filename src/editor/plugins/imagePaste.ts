import { isHistoryTransaction } from "prosemirror-history";
import { Plugin, PluginKey, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { replacementShut } from "../../schema/locks";
import { editsShut } from "../../schema/protectionState";
import {
  hasResolvableClipboardImages,
  type ResolvedClipboardImages,
  resolveClipboardImages,
} from "../clipboard/images";
import { documentBodyHeightPx, documentBodyWidthPx } from "../documentStyles";
import { richHtmlSlice } from "../externalClipboard";
import { insertPlainTextAt } from "../plainText";

type PasteId = object;

interface ImagePasteMeta {
  add?: { id: PasteId; from: number; to: number };
  remove?: readonly PasteId[];
}

interface ImagePasteState {
  decorations: DecorationSet;
  cancelVersion: number;
}

interface PasteTask {
  id: PasteId;
  view: EditorView;
  html: string;
  text: string;
  maxWidthPx: number;
  maxHeightPx: number;
  canceled: boolean;
  controller: AbortController | null;
}

const LOAD_TIMEOUT_MS = 15_000;
const imagePasteKey = new PluginKey<ImagePasteState>("docxImagePaste");

function marker(view: EditorView): HTMLElement {
  const element = view.dom.ownerDocument.createElement("span");
  element.hidden = true;
  return element;
}

function taskRange(
  decorations: DecorationSet,
  id: PasteId
): { from: number; to: number } | null {
  const found = decorations.find(
    undefined,
    undefined,
    (spec) => spec.task === id
  );
  const from = found.find((decoration) => decoration.spec.edge === "from");
  const to = found.find((decoration) => decoration.spec.edge === "to");
  return from && to ? { from: from.from, to: to.from } : null;
}

function rangesConflict(
  first: { from: number; to: number },
  second: { from: number; to: number }
): boolean {
  if (first.from === first.to) {
    return second.from === second.to
      ? first.from === second.from
      : second.from <= first.from && first.from < second.to;
  }
  if (second.from === second.to) {
    return first.from <= second.from && second.from < first.to;
  }
  return first.from < second.to && second.from < first.to;
}

function transactionTouchesRange(
  transaction: Transaction,
  initial: { from: number; to: number }
): boolean {
  let { from, to } = initial;
  for (const [index, step] of transaction.steps.entries()) {
    const map = transaction.mapping.maps[index];
    if (!map) return true;
    let described = false;
    let touched = false;
    map.forEach((oldStart, oldEnd) => {
      described = true;
      if (
        oldStart === oldEnd
          ? from <= oldStart && oldStart < to
          : oldStart < to && from < oldEnd
      ) {
        touched = true;
      }
    });
    if (touched) return true;
    if (!described) {
      const json = step.toJSON() as { from?: unknown; to?: unknown };
      if (typeof json.from !== "number" || typeof json.to !== "number") {
        return true;
      }
      if (json.from < to && from < json.to) return true;
    }
    from = map.map(from, -1);
    to = map.map(to, -1);
  }
  return false;
}

export function imagePaste(): Plugin<ImagePasteState> {
  // Markers map a pending paste through unrelated edits. Any edit inside its selected range,
  // history move, or overlapping paste removes the markers before completion can overwrite it.
  const tasks: PasteTask[] = [];
  let working = false;

  const cancel = (task: PasteTask) => {
    task.canceled = true;
    task.controller?.abort();
  };

  const cancelAll = () => {
    for (const task of tasks) cancel(task);
  };

  const removeMarker = (task: PasteTask) => {
    if (task.view.isDestroyed) return;
    task.view.dispatch(
      task.view.state.tr
        .setMeta(imagePasteKey, {
          remove: [task.id],
        } satisfies ImagePasteMeta)
        .setMeta("addToHistory", false)
    );
  };

  const finish = (
    task: PasteTask,
    resolved: ResolvedClipboardImages | null
  ) => {
    if (task.canceled || task.view.isDestroyed) return;
    const state = imagePasteKey.getState(task.view.state);
    const range = state && taskRange(state.decorations, task.id);
    if (!range) return;
    removeMarker(task);
    if (task.view.isDestroyed) return;
    const slice =
      resolved === null || resolved.images.size === 0
        ? null
        : richHtmlSlice(
            task.view.state,
            task.view.dom.ownerDocument,
            resolved.source,
            resolved.images
          );
    if (slice !== null) {
      task.view.dispatch(
        task.view.state.tr
          .replaceRange(range.from, range.to, slice)
          .scrollIntoView()
      );
    } else {
      insertPlainTextAt(task.view, task.text, range.from, range.to);
    }
  };

  const work = async () => {
    if (working) return;
    working = true;
    try {
      while (tasks.length > 0) {
        const task = tasks[0];
        if (!task || task.canceled || task.view.isDestroyed) {
          tasks.shift();
          continue;
        }
        const controller = new AbortController();
        task.controller = controller;
        const timeout = setTimeout(() => controller.abort(), LOAD_TIMEOUT_MS);
        let resolved: ResolvedClipboardImages | null = null;
        try {
          const request = resolveClipboardImages(
            task.view.dom.ownerDocument,
            task.html,
            {
              maxWidthPx: task.maxWidthPx,
              maxHeightPx: task.maxHeightPx,
              signal: controller.signal,
            }
          );
          resolved = request === null ? null : await request;
        } catch {
          resolved = null;
        } finally {
          clearTimeout(timeout);
          task.controller = null;
        }
        finish(task, resolved);
        tasks.shift();
      }
    } finally {
      working = false;
    }
  };

  return new Plugin({
    key: imagePasteKey,
    state: {
      init: (): ImagePasteState => ({
        decorations: DecorationSet.empty,
        cancelVersion: 0,
      }),
      apply(transaction, current) {
        if (isHistoryTransaction(transaction)) {
          return {
            decorations: DecorationSet.empty,
            cancelVersion: current.cancelVersion + 1,
          };
        }
        let decorations = current.decorations;
        if (transaction.docChanged) {
          const ids = new Set<PasteId>();
          for (const decoration of decorations.find()) {
            ids.add(decoration.spec.task as PasteId);
          }
          const invalidated = new Set<PasteId>();
          for (const id of ids) {
            const range = taskRange(decorations, id);
            if (
              range !== null &&
              range.from !== range.to &&
              transactionTouchesRange(transaction, range)
            ) {
              invalidated.add(id);
            }
          }
          decorations = decorations.remove(
            decorations.find(undefined, undefined, (spec) =>
              invalidated.has(spec.task as PasteId)
            )
          );
        }
        decorations = decorations.map(transaction.mapping, transaction.doc);
        const meta = transaction.getMeta(imagePasteKey) as
          | ImagePasteMeta
          | undefined;
        if (meta?.remove) {
          const removed = new Set(meta.remove);
          decorations = decorations.remove(
            decorations.find(undefined, undefined, (spec) =>
              removed.has(spec.task as PasteId)
            )
          );
        }
        if (meta?.add) {
          const { id, from, to } = meta.add;
          decorations = decorations.add(transaction.doc, [
            Decoration.widget(from, marker, {
              task: id,
              edge: "from",
              side: -1,
            }),
            Decoration.widget(to, marker, {
              task: id,
              edge: "to",
              side: -1,
            }),
          ]);
        }
        return { decorations, cancelVersion: current.cancelVersion };
      },
    },
    view: () => ({
      update(view, previous) {
        const before = imagePasteKey.getState(previous);
        const current = imagePasteKey.getState(view.state);
        if (!current) return;
        if (before && before.cancelVersion !== current.cancelVersion) {
          cancelAll();
          return;
        }
        for (const task of tasks) {
          if (!taskRange(current.decorations, task.id)) cancel(task);
        }
      },
      destroy: cancelAll,
    }),
    props: {
      decorations: (state) => imagePasteKey.getState(state)?.decorations,
      handlePaste(view, event) {
        if (
          editsShut(view.state) ||
          replacementShut(view.state.selection, view.state.doc)
        ) {
          return false;
        }
        if (view.state.selection.ranges.length > 1) return false;
        const html = event.clipboardData?.getData("text/html") ?? "";
        if (!hasResolvableClipboardImages(view.dom.ownerDocument, html)) {
          return false;
        }
        const text = event.clipboardData?.getData("text/plain") ?? "";
        const { from, to } = view.state.selection;
        const current = imagePasteKey.getState(view.state);
        const replaced = tasks.filter((task) => {
          if (!current || task.canceled) return false;
          const range = taskRange(current.decorations, task.id);
          return range !== null && rangesConflict(range, { from, to });
        });
        for (const task of replaced) cancel(task);
        const id: PasteId = {};
        view.dispatch(
          view.state.tr
            .setMeta(imagePasteKey, {
              add: { id, from, to },
              remove: replaced.map((task) => task.id),
            } satisfies ImagePasteMeta)
            .setMeta("addToHistory", false)
        );
        tasks.push({
          id,
          view,
          html,
          text,
          maxWidthPx: documentBodyWidthPx(view.state),
          maxHeightPx: documentBodyHeightPx(view.state),
          canceled: false,
          controller: null,
        });
        void work();
        return true;
      },
    },
  });
}
