/** Registers every command module. */
import './file';
import './edit';
import './format';
import './cells';
import './view';
import.meta.glob(['./*.ts', '!./index.ts', '!./registry.ts', '!./ops.ts'], { eager: true });
export { def, getCommand, allCommands, commandCount, searchCommands } from './registry';

// --------------------------------------------------------------------------------------------
// Tools that write straight to cells or styles, without the per-cell checks the editing commands
// make: on a protected sheet they are refused outright (unlocking cells would otherwise defeat
// the protection without its password).
import { getCommand as lookup } from './registry';
import { toast as say } from '../ui/dialog';

const SHEET_GUARDED = ['util.unlockSelection', 'util.lockSelection', 'util.lockNonEmpty', 'util.lockFormulas', 'plus.batchCheckboxes', 'plus.insertRandomPassword', 'util.formatByContent', 'util.colorByContentType', 'util.highlightCells'];
const BOOK_GUARDED = ['util.trimWorkbook', 'util.formulasToValuesBook', 'util.fontAllSheets'];
for (const id of [...SHEET_GUARDED, ...BOOK_GUARDED]) {
  const c = lookup(id);
  if (!c) continue;
  const run = c.run;
  const wholeBook = BOOK_GUARDED.includes(id);
  c.run = (app, arg) => {
    const blocked = wholeBook ? app.book.sheets.find((s) => s.protection.on) : app.sheet.protection.on ? app.sheet : undefined;
    if (blocked) { say(`The sheet "${blocked.name}" is protected. Unprotect it (Review ▸ Protect Sheet) to make this change.`); return false; }
    return run(app, arg);
  };
}
