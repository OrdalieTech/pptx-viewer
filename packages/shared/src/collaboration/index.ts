export * from '../render/collaboration-sync';
export * from '../render/collaboration-reconcile';
export { findElementYMap } from '../render/collaboration-live-patch-target';
// Merge-aware row and column edits of a table model, for headless callers of the codec.
export {
	insertTableRow,
	deleteTableRow,
	insertTableColumn,
	deleteTableColumn,
} from '../render/table-layout';
