import { removeAllTags, removeTagsFromWorksheet, removeTagsFromSelection } from "./excel";

/* global Office */

Office.onReady(() => {
  Office.actions.associate("removeAllTags", removeAllTags);
  Office.actions.associate("removeTagsFromWorksheet", removeTagsFromWorksheet);
  Office.actions.associate("removeTagsFromSelection", removeTagsFromSelection);
});
