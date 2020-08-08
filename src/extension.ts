import * as vscode from 'vscode';
import { PowerQueryMCodeReader, ExcelRegistry } from "./ExcelHandler";


// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
	let logger = function(msg: string) {vscode.window.showInformationMessage("EditExcelPQM: " + msg);};
	let errorLog = function(msg: string) {vscode.window.showErrorMessage("EditExcelPQM: " + msg);};
	let excelRegistry = new ExcelRegistry(logger);

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
