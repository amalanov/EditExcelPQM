import * as vscode from 'vscode';
import { activate_winax, PowerQueryMCodeReader, ExcelRegistry } from "./ExcelHandler";


// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
	let logger = function(msg: string) {vscode.window.showInformationMessage("EditExcelPQM: " + msg);};
	let errorLog = function(msg: string) {vscode.window.showErrorMessage("EditExcelPQM: " + msg);};
	let excelRegistry = new ExcelRegistry(logger);

	try{
		activate_winax();
	} catch (e){
		errorLog("Hi! This is Sasha. Erro occured on load of winax module. " +
				 "This is native node module. This could happen due to update of " +
				 "VSCode Electron version. In this case you need either to download a new " +
				 "version from my repo or change electron version in nmp task build_winax_for_vscode " +
				 "and then run in Windows command line \n" +
				 "nmp build_vscode_extension && nmp pack_extension\n" +
				 "so the error is....\n" + e.message);
		return;
	}

	let extract = vscode.commands.registerCommand('extension.extract_pqm_from_excel', (file: vscode.Uri) => {
		try{
			logger("Start export Excel->M");
			let loader = new PowerQueryMCodeReader(file.fsPath, excelRegistry);
			loader.importFromExcel();
			loader.exportToFile();
			logger("Finish export Excel->M");
		} catch (e){
			if (e instanceof Error){
				errorLog("Error while extracting from Excel: " + e.message);
			}
		}
	});
	context.subscriptions.push(extract);

	let save = vscode.commands.registerCommand('extension.extract_pqm_to_excel', (file: vscode.Uri) => {
		try{
			logger("Start export M->Excel");
			let loader = new PowerQueryMCodeReader(file.fsPath, excelRegistry);
			loader.importFromFile();
			loader.exportToExcel();
			logger("Finish export M->Excel");
		} catch (e) {
			if (e instanceof  Error){
				errorLog("EditExcelPQM: Error while saving to Excel: " + e.message);
			}
		}
	});
	context.subscriptions.push(save);

	let closeSave = vscode.commands.registerCommand('extension.close_excel_save', (file: vscode.Uri) => {
		try{
			logger("Save changes and close xlsx file");
			excelRegistry.close(file.fsPath, true);
			logger("Finish saving and closing");
		} catch (e) {
			if (e instanceof  Error){
				errorLog("EditExcelPQM: Error while closing Excel: " + e.message);
			}
		}
	});
	context.subscriptions.push(closeSave);

	let closeNoSave = vscode.commands.registerCommand('extension.close_excel_nosave', (file: vscode.Uri) => {
		try{
			logger("Close xlsx file without saving");
			excelRegistry.close(file.fsPath, false);
			logger("Finish closing");
		} catch (e) {
			if (e instanceof  Error){
				errorLog("Error while closing Excel: " + e.message);
			}
		}
	});
	context.subscriptions.push(closeNoSave);

	context.subscriptions.push(excelRegistry);
}

  
export function deactivate() {}
