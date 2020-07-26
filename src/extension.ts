import * as vscode from 'vscode';
import { PowerQueryMCodeReader, ExcelRegistry } from "./ExcelHandler";


// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
	let excelRegistry = new ExcelRegistry(function(msg: string) {vscode.window.showInformationMessage(msg)});

	let extract = vscode.commands.registerCommand('extension.extract_pqm_from_excel', (file: vscode.Uri) => {
		try{
			vscode.window.showInformationMessage("Start export Excel->M");
			let loader = new PowerQueryMCodeReader(file.fsPath, excelRegistry);
			loader.importFromExcel();
			loader.exportToFile();
			vscode.window.showInformationMessage("Finish export Excel->M");
		} catch (e){
			if (e instanceof Error){
				vscode.window.showErrorMessage("Error while extracting from Excel: " + e.message);
			}
		}
	});
	context.subscriptions.push(extract);

	let save = vscode.commands.registerCommand('extension.extract_pqm_to_excel', (file: vscode.Uri) => {
		try{
			vscode.window.showInformationMessage("Start export M->Excel");
			let loader = new PowerQueryMCodeReader(file.fsPath, excelRegistry);
			loader.importFromFile();
			loader.exportToExcel();
			vscode.window.showInformationMessage("Finish export M->Excel");
		} catch (e) {
			if (e instanceof  Error){
				vscode.window.showErrorMessage("Error while saving to Excel: " + e.message);
			}
		}
	});
	context.subscriptions.push(save);

	let close = vscode.commands.registerCommand('extension.close_excel', (file: vscode.Uri) => {
		try{
			vscode.window.showInformationMessage("Save changes and close xlsx file");
			excelRegistry.close(file.fsPath);
			vscode.window.showInformationMessage("Finish saving and closing");
		} catch (e) {
			if (e instanceof  Error){
				vscode.window.showErrorMessage("Error while closing Excel: " + e.message);
			}
		}
	});
	context.subscriptions.push(close);

	context.subscriptions.push(excelRegistry);
}

  
export function deactivate() {}
