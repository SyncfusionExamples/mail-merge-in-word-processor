import { createElement, Event, KeyboardEventArgs } from '@syncfusion/ej2-base';
import { DocumentEditor,DocumentEditorContainer, FormatType } from '@syncfusion/ej2-angular-documenteditor';
import { Button } from '@syncfusion/ej2-angular-buttons';
import { DropDownButton, ItemModel } from '@syncfusion/ej2-angular-splitbuttons';

import { MenuEventArgs } from '@syncfusion/ej2-angular-navigations';
/**
 * Represents document editor title bar.
 */
export class TitleBar {
    private tileBarDiv: HTMLElement;
    private documentTitle: HTMLElement;
    private documentTitleContentEditor: HTMLElement;
    private export: DropDownButton;
    private print: Button;
    private open: Button;
    private documentEditor: DocumentEditor;
    private isRtl: boolean;
    private previousBtn;
    private nextBtn;
    private checkBox:HTMLInputElement;
    private innerDiv;
    public currentRecordId = 0;
    public recordCount = 1;
    public isOpenInternal = false;
    private container:DocumentEditorContainer;
        
    public isMergePreview = false;
    public mergeTemplateDocxBlob = null;
    constructor(element: HTMLElement, docEditor: DocumentEditor,container:DocumentEditorContainer, isShareNeeded: Boolean, isRtl?: boolean) {
        this.isRtl = isRtl;
        //initializes title bar elements.
        this.tileBarDiv = element;
        this.container=container;
        this.documentEditor = docEditor;
        this.initializeTitleBar(isShareNeeded);
        this.wireEvents();
    }
    private initializeTitleBar = (isShareNeeded: Boolean): void => {
        let downloadText: string;
        let downloadToolTip: string;
        let printText: string;
        let printToolTip: string;
        let openText: string;
        let documentTileText: string;
        if (!this.isRtl) {
            downloadText = 'Download';
            downloadToolTip = 'Download this document.';
            printText = 'Print';
            printToolTip = 'Print this document (Ctrl+P).';
            openText = 'Open';
            documentTileText = 'Document Name. Click or tap to rename this document.';
        } else {
            downloadText = 'تحميل';
            downloadToolTip = 'تحميل هذا المستند';
            printText = 'طباعه';
            printToolTip = 'طباعه هذا المستند (Ctrl + P)';
            openText = 'فتح';
            documentTileText = 'اسم المستند. انقر أو اضغط لأعاده تسميه هذا المستند';
        }
        // tslint:disable-next-line:max-line-length
        this.documentTitle = createElement('label', { id: 'documenteditor_title_name', styles: 'font-weight:400;text-overflow:ellipsis;white-space:pre;overflow:hidden;user-select:none;cursor:text' });
        let iconCss: string = 'e-de-padding-right';
        let btnFloatStyle: string = 'float:right;';
        let titleCss: string = '';
        if (this.isRtl) {
            iconCss = 'e-de-padding-right-rtl';
            btnFloatStyle = 'float:left;';
            titleCss = 'float:right;';
        }
        // tslint:disable-next-line:max-line-length
        this.documentTitleContentEditor = createElement('div', { id: 'documenteditor_title_contentEditor', className: 'single-line', styles: titleCss });
        this.documentTitleContentEditor.appendChild(this.documentTitle);
        this.tileBarDiv.appendChild(this.documentTitleContentEditor);
        this.documentTitleContentEditor.setAttribute('title', 'Document Name. Click or tap to rename this document.');
        let btnStyles: string = btnFloatStyle + 'background: transparent;box-shadow:none; font-family: inherit;border-color: transparent;'
            + 'border-radius: 2px;color:inherit;font-size:12px;text-transform:capitalize;margin-top:4px;height:28px;font-weight:400;'
            + 'margin-top: 2px;';
        // tslint:disable-next-line:max-line-length
        this.print = this.addButton('e-de-icon-Print ' + iconCss, printText, btnStyles, 'de-print', printToolTip, false) as Button;
        this.open = this.addButton('e-de-icon-Open ' + iconCss, openText, btnStyles, 'de-open', documentTileText, false) as Button;
        let items: ItemModel[] = [
            { text: 'Microsoft Word (.docx)', id: 'word' },
            { text: 'Syncfusion Document Text (.sfdt)', id: 'sfdt' },
        ];
        // tslint:disable-next-line:max-line-length
        this.export = this.addButton('e-de-icon-Download ' + iconCss, downloadText, btnStyles, 'documenteditor-share', downloadToolTip, true, items) as DropDownButton;
        if (!isShareNeeded) {
            this.export.element.style.display = 'none';
        } else {
            this.open.element.style.display = 'none';
        }
        var divMarginStyle = 'margin:5px';
        var mailmergeDivStyle = 'float:right;display:inline-flex;display:none';
        var mailmergeDiv = createElement('div', { id: 'mailmergeDiv', styles: mailmergeDivStyle });
    
        this.tileBarDiv.appendChild(mailmergeDiv);
        this.previousBtn = createElement('button', { id: 'previousBtn' });
        this.previousBtn.innerHTML = '&laquo;';
        mailmergeDiv.appendChild(this.previousBtn);
        this.innerDiv = createElement('div', { id: 'innerDiv', styles: divMarginStyle });
        this.innerDiv.innerText = 'Record ' + (this.currentRecordId + 1).toString();
        mailmergeDiv.appendChild(this.innerDiv);
        this.nextBtn = createElement('button', { id: 'nextBtn' });
        this.nextBtn.innerHTML = '&raquo;';
        mailmergeDiv.appendChild(this.nextBtn);
        var exitBtn = createElement('button', { id: 'exitBtn' });
        exitBtn.innerHTML = 'Exit preview';
        exitBtn.setAttribute('title', 'Exit the Mail merge preview window.');
        mailmergeDiv.appendChild(exitBtn);
        this.checkBox = createElement('input', { id: 'checkBox'}) as HTMLInputElement;
        this.checkBox.type = "checkbox";
        this.checkBox.name = "name";
        mailmergeDiv.appendChild(this.checkBox);
        var checkBoxDiv = createElement('div', { id: 'checkBoxDiv', styles: divMarginStyle });
        checkBoxDiv.innerText = 'Download as individual files';
        checkBoxDiv.setAttribute('title', 'Download as individual files for all records.');
        mailmergeDiv.appendChild(checkBoxDiv);
        this.previousBtn.addEventListener('click', this.onPreviousBtn.bind(this));
        this.nextBtn.addEventListener('click', this.onNextBtn.bind(this));
        exitBtn.addEventListener('click', this.onExitBtn.bind(this));
    }
    onPreviousBtn() {
        debugger;
        if (this.currentRecordId > 0) {
            this.currentRecordId--;
            this.innerDiv.innerText = 'Record ' + (this.currentRecordId + 1).toString();
            this.previewCurrentRecord();
        }
    }
    onNextBtn() {
        debugger;
        if (this.currentRecordId < this.recordCount - 1) {
            this.currentRecordId++;
            this.innerDiv.innerText = 'Record ' + (this.currentRecordId + 1).toString();
            this.previewCurrentRecord();
        }
    }
    previewCurrentRecord() {
        debugger;
        var sfdtFiles=JSON.parse(sessionStorage.getItem("sfdtFiles"));
        if (sfdtFiles != null && sfdtFiles.length > 0 && this.currentRecordId < sfdtFiles.length) {
          this.isOpenInternal = true;
          this.documentEditor.open(sfdtFiles[this.currentRecordId]);
          this.isOpenInternal = false;
            //Hides the list of fields, while displaying the final report.
            this.hideFieldList();
        }
      }
      
showFieldList() {
    document.getElementById('fieldDiv').style.display = '';
    this.enableMergeOption();
  }
  hideFieldList() {
    document.getElementById('fieldDiv').style.display = 'none';
    this.disableMergedOption();
  }
  enableMergeOption() {
    this.container.toolbar.enableItems(7, true);
    //container.toolbar.enableItems(7, true);
  }
  disableMergedOption() {
    this.container.toolbar.enableItems(7, false);
    //container.toolbar.enableItems(7, false)
  }
  showHideWaitingIndicator(show: boolean): void {
    let waitingPopUp: HTMLElement = document.getElementById('waiting-popup');
    let inActiveDiv:HTMLElement = document.getElementById('popup-overlay');
    inActiveDiv.style.display = show ? 'block' : 'none';
    waitingPopUp.style.display = show ? 'block' : 'none';
}

showMergePreviewOption() {
    this.isMergePreview = true;
    var mailmergeDiv = document.getElementById('mailmergeDiv');
    if (mailmergeDiv != null) {
        mailmergeDiv.style.display = 'inline-flex';
    }
  }
  hideMergePreviewOption() {
    this.isMergePreview = false;
    this.currentRecordId = 0;
    //Clears the SFDT files from cache.
    sessionStorage.setItem("sfdtFiles", "");
    var mailmergeDiv = document.getElementById('mailmergeDiv');
    if (mailmergeDiv != null) {
        mailmergeDiv.style.display = 'none';
    }
  }
    onExitBtn() {
        debugger;
        var obj=this;
        this.showHideWaitingIndicator(true)
        var httpRequest = new XMLHttpRequest();
        httpRequest.open('Post', 'api/DocumentEditor/Import', true);
        httpRequest.onreadystatechange = function () {
            if (httpRequest.readyState === 4) {
                if (httpRequest.status === 200 || httpRequest.status === 304) {
                    obj.isOpenInternal = true;
                    obj.documentEditor.open(httpRequest.responseText);
                    obj.isOpenInternal = false;
                    obj.mergeTemplateDocxBlob = null;
                }
            }
            obj.showHideWaitingIndicator(false);
        };
        var formData = new FormData();
        formData.append('files', obj.mergeTemplateDocxBlob, obj.mergeTemplateDocxBlob.name);
        httpRequest.send(formData);
        this.showFieldList();
        obj.hideMergePreviewOption();
    }
    private setTooltipForPopup(): void {
        // tslint:disable-next-line:max-line-length
        document.getElementById('documenteditor-share-popup').querySelectorAll('li')[0].setAttribute('title', 'Download a copy of this document to your computer as a DOCX file.');
        // tslint:disable-next-line:max-line-length
        document.getElementById('documenteditor-share-popup').querySelectorAll('li')[1].setAttribute('title', 'Download a copy of this document to your computer as an SFDT file.');
    }
    private wireEvents = (): void => {
        this.print.element.addEventListener('click', this.onPrint);
        this.open.element.addEventListener('click', (e: Event) => {
            if ((e.target as HTMLInputElement).id === 'de-open') {
                let fileUpload: HTMLInputElement = document.getElementById('uploadfileButton') as HTMLInputElement;
                fileUpload.value = '';
                fileUpload.click();
            }
        });
        this.documentTitleContentEditor.addEventListener('keydown', (e: KeyboardEventArgs) => {
            if (e.keyCode === 13) {
                e.preventDefault();
                this.documentTitleContentEditor.contentEditable = 'false';
                if (this.documentTitleContentEditor.textContent === '') {
                    this.documentTitleContentEditor.textContent = 'Document1';
                }
            }
        });
        this.documentTitleContentEditor.addEventListener('blur', (): void => {
            if (this.documentTitleContentEditor.textContent === '') {
                this.documentTitleContentEditor.textContent = 'Document1';
            }
            this.documentTitleContentEditor.contentEditable = 'false';
            this.documentEditor.documentName = this.documentTitle.textContent;
        });
        this.documentTitleContentEditor.addEventListener('click', (): void => {
            this.updateDocumentEditorTitle();
        });
    }
    private updateDocumentEditorTitle = (): void => {
        this.documentTitleContentEditor.contentEditable = 'true';
        this.documentTitleContentEditor.focus();
        window.getSelection().selectAllChildren(this.documentTitleContentEditor);
    }
    // Updates document title.
    public updateDocumentTitle = (): void => {
        if (this.documentEditor.documentName === '') {
            this.documentEditor.documentName = 'Untitled';
        }
        this.documentTitle.textContent = this.documentEditor.documentName;
    }
    public getHeight(): number {
        return this.tileBarDiv.offsetHeight + 4;
    }
    // tslint:disable-next-line:max-line-length
    private addButton(iconClass: string, btnText: string, styles: string, id: string, tooltipText: string, isDropDown: boolean, items?: ItemModel[]): Button | DropDownButton {
        let button: HTMLButtonElement = createElement('button', { id: id, styles: styles }) as HTMLButtonElement;
        this.tileBarDiv.appendChild(button);
        button.setAttribute('title', tooltipText);
        if (isDropDown) {
            // tslint:disable-next-line:max-line-length
            let dropButton: DropDownButton = new DropDownButton({ select: this.onExportClick, items: items, iconCss: iconClass, cssClass: 'e-caret-hide', content: btnText, open: (): void => { this.setTooltipForPopup(); } }, button);
            return dropButton;
        } else {
            let ejButton: Button = new Button({ iconCss: iconClass, content: btnText }, button);
            return ejButton;
        }
    }
    private onPrint = (): void => {
        this.documentEditor.print();
    }
    private onExportClick = (args: MenuEventArgs): void => {
        let value: string = args.item.id;
        switch (value) {
            case 'word':
                this.save('Docx');
                break;
            case 'sfdt':
                this.save('Sfdt');
                break;
        }
    }
    private save = (format: string): void => {
        // tslint:disable-next-line:max-line-length
        this.documentEditor.save(this.documentEditor.documentName === '' ? 'example' : this.documentEditor.documentName, format as FormatType);
    }
}