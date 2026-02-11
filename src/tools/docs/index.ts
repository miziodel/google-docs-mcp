import type { FastMCP } from 'fastmcp';

// Core read/write
import { register as readGoogleDoc } from './readGoogleDoc.js';
import { register as listDocumentTabs } from './listDocumentTabs.js';
import { register as appendToGoogleDoc } from './appendToGoogleDoc.js';
import { register as insertText } from './insertText.js';
import { register as deleteRange } from './deleteRange.js';

// Formatting
import { register as applyTextStyle } from './applyTextStyle.js';
import { register as applyParagraphStyle } from './applyParagraphStyle.js';
import { register as formatMatchingText } from './formatMatchingText.js';

// Structure
import { register as insertTable } from './insertTable.js';
import { register as editTableCell } from './editTableCell.js';
import { register as insertPageBreak } from './insertPageBreak.js';
import { register as insertImageFromUrl } from './insertImageFromUrl.js';
import { register as insertLocalImage } from './insertLocalImage.js';
import { register as fixListFormatting } from './fixListFormatting.js';
import { register as findElement } from './findElement.js';

// Comments
import { register as listComments } from './listComments.js';
import { register as getComment } from './getComment.js';
import { register as addComment } from './addComment.js';
import { register as replyToComment } from './replyToComment.js';
import { register as resolveComment } from './resolveComment.js';
import { register as deleteComment } from './deleteComment.js';

export function registerDocsTools(server: FastMCP) {
  // Core read/write
  readGoogleDoc(server);
  listDocumentTabs(server);
  appendToGoogleDoc(server);
  insertText(server);
  deleteRange(server);

  // Formatting
  applyTextStyle(server);
  applyParagraphStyle(server);
  formatMatchingText(server);

  // Structure
  insertTable(server);
  editTableCell(server);
  insertPageBreak(server);
  insertImageFromUrl(server);
  insertLocalImage(server);
  fixListFormatting(server);
  findElement(server);

  // Comments
  listComments(server);
  getComment(server);
  addComment(server);
  replyToComment(server);
  resolveComment(server);
  deleteComment(server);
}
