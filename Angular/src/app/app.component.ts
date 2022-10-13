import { Component, ViewChild } from '@angular/core';
import { ToolbarService, DocumentEditorContainerComponent } from '@syncfusion/ej2-angular-documenteditor';
import { TitleBar } from './title-bar';
import * as jsonData from './Employees.json';
import { ClickEventArgs } from '@syncfusion/ej2-navigations/src/toolbar/toolbar';
import { ListView, SelectEventArgs } from '@syncfusion/ej2-lists';
import { isNullOrUndefined } from '@syncfusion/ej2-base';
import { Dialog, DialogUtility } from '@syncfusion/ej2-popups';
import { DropDownButton } from '@syncfusion/ej2-splitbuttons';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  providers: [ToolbarService]
})
export class AppComponent {
  //Gets the DocumentEditorContainerComponent instance from view DOM using a template reference variable 'documenteditor_ref'.
  @ViewChild('documenteditor_ref') public container! : DocumentEditorContainerComponent;
  public serviceLink: string;
  titleBar: TitleBar;
  listview;
    dialogBox;
    jsonDataSource;
    tempListData = [];

   
    drpDownBtn=null;
  public listData: { [key: string]: Object }[] = [
    {
        text: 'ProductName',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'ShipName',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'CustomerID',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'Quantity',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'UnitPrice',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'Discount',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'ShipAddress',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'ShipCity',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'ShipCountry',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'OrderId',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    },
    {
        text: 'OrderDate',
        category: 'Drag or click the field to insert.',
        htmlAttributes: { draggable: true }
    }
];
public item =['New', 'Open', 'Separator', 'Undo',
'Redo',
'Separator',
{
    prefixIcon: 'sf-icon-InsertMergeField',
    tooltipText: 'Insert Field',
    text: this.onWrapText('Insert Field'),
    id: 'InsertField'
},
{
    prefixIcon: 'sf-icon-FinishMerge',
    tooltipText: 'Merge Document',
    text: this.onWrapText('Merge Document'),
    id: 'MergeDocument'
},
'Separator',
'Image',
'Table',
'Hyperlink',
'Bookmark',
'TableOfContents',
'Separator',
'Header',
'Footer',
'PageSetup',
'PageNumber',
'Break',
'Separator',
'Find',
'Separator',
'Comments',
'TrackChanges',
'Separator',
'LocalClipboard',
'RestrictEditing',
'Separator',
'FormFields',
'UpdateFields'
];
onWrapText(text: string): string {
let content: string = '';
let index: number = text.lastIndexOf(' ');
content = text.slice(0, index);
text.slice(index);
content += '<div class="e-de-text-wrap">' + text.slice(index) + '</div>';
return content;
}
toolbarClick = (args: ClickEventArgs): void => {
  switch (args.item.id) {
      case 'InsertField':
          this.showInsertFielddialog();
          break;
  }
};
  constructor() {
    this.serviceLink = 'http://localhost:62869/api/documenteditor/';
  }
  onCreate(): void {
  
    let titleBarElement: HTMLElement = document.getElementById('default_title_bar');
    this.titleBar = new TitleBar(titleBarElement, this.container.documentEditor,this.container, true);
    //Opens the default template Getting Started.docx from web API.
    this.openTemplate();
    this.container.documentEditor.documentName = 'Getting Started';

    this.titleBar.updateDocumentTitle();
     //Initialize action items.
     var items = [
      {
          text: 'Preview Results',
          id: 'preview'
      },
      {
          text: 'Merge & Finish',
          id: 'Merge'
      }];
  if (this.drpDownBtn == null || this.drpDownBtn == undefined) {
    this.drpDownBtn = new DropDownButton({ items: items, cssClass: 'e-caret-hide', select: this.onSelectMergeItems.bind(this) });

      // Render initialized DropDownButton.
      this.drpDownBtn.appendTo('#MergeDocument');
  }

   document.getElementById("listview").addEventListener("dragstart", function (event) {
      event.dataTransfer.setData("Text", (event.target as any).innerText);
      (event.target as any).classList.add('de-drag-target');

    });

    // Prevent default drag over for document editor element
    this.container.documentEditor.element.addEventListener("dragover", function (event) {
      event.preventDefault();
    });

    // Drop Event for document editor element
    this.container.documentEditor.element.addEventListener("drop", (e) => {
      var text = e.dataTransfer.getData('Text');//.replace(/\n/g, '').replace(/\r/g, '').replace(/\r\n/g, '');
      this.container.documentEditor.selection.select({ x: e.offsetX, y: e.offsetY, extend: false });
      //this.container.documentEditor.editor.insertText(text);
      this.insertField(text);
    });

    document.addEventListener("dragend", (event) => {
      if ((event.target as any).classList.contains('de-drag-target')) {
        (event.target as any).classList.remove('de-drag-target');
      }
    });
    setInterval(()=>{
      this.updateDocumentEditorSize();
    }, 100);
    //Adds event listener for browser window resize event.
    window.addEventListener("resize", this.onWindowResize);
    debugger;
    let data=jsonData;
    this.loadJSON(JSON.stringify((data as any).default));
  }
  onSelectMergeItems(args) {
    debugger;
    switch (args.item.id) {
        case 'preview':
            this.titleBar.showMergePreviewOption();
            break;
    }
    var obj=this;
    if (this.titleBar.mergeTemplateDocxBlob == null) {
        this.container.documentEditor.saveAsBlob('Docx').then(function (exportedDocument) {
          obj.titleBar.mergeTemplateDocxBlob = exportedDocument;
          obj.titleBar.mergeTemplateDocxBlob.name = obj.container.documentEditor.documentName + '.docx';
          obj.mergeDocument(obj.titleBar.mergeTemplateDocxBlob);
        });
    }
    else {
      this.mergeDocument(this.titleBar.mergeTemplateDocxBlob);
    }
}
previewCurrentRecord() {
  var sfdtFiles=JSON.parse(sessionStorage.getItem("sfdtFiles"));
  if (sfdtFiles != null && sfdtFiles.length > 0 && this.titleBar.currentRecordId < sfdtFiles.length) {
    this.titleBar.isOpenInternal = true;
    this.container.documentEditor.open(sfdtFiles[this.titleBar.currentRecordId]);
    this.titleBar.isOpenInternal = false;
      //Hides the list of fields, while displaying the final report.
      this.titleBar.hideFieldList();
  }
}
  onDocumentChange(): void {
    if (!isNullOrUndefined(this.titleBar)) {
        this.titleBar.updateDocumentTitle();
    }
    if (this.titleBar&&this.titleBar.isOpenInternal == false) {
      //Shows the list of fields, while displaying the template.
      this.titleBar.showFieldList();
      this.titleBar.hideMergePreviewOption();
      this.titleBar.mergeTemplateDocxBlob = null;
  }
    this.container.documentEditor.focusIn();
  }

  onDestroy(): void {
    //Removes event listener for browser window resize event.
    window.removeEventListener("resize", this.onWindowResize);
  }

  onWindowResize= (): void => {
    //Resizes the document editor component to fit full browser window automatically whenever the browser resized.
    this.updateDocumentEditorSize();
  }

  updateDocumentEditorSize(): void {
     //Resizes the document editor component to fit full browser window.
     var windowWidth = document.getElementById('default_title_bar').offsetWidth - (document.getElementById('fieldDiv').offsetWidth);
     var windowHeight = window.innerHeight - (document.getElementById('default_title_bar').offsetHeight);
     this.container.resize(windowWidth, windowHeight);
  }
  mergeDocument(blob:any) {
    let obj=this;
    let sfdtFiles = [];
    let jsonObject:any = {};
    jsonObject.MergeSettings = [];
    let mergeSetting:any = {};
    mergeSetting.ClearFields = true;
    mergeSetting.RemoveEmptyParagraphs = false;
    mergeSetting.RemoveEmptyGroup = false;
    mergeSetting.InsertAsNewRow = false;
    mergeSetting.StartNextRecordAtNewPage = true;
    mergeSetting.MergeType = "NestedGroup";
    mergeSetting.MergeCommand = "Employees";
    var recordCnt = "";
    if (this.titleBar.isMergePreview) {
        recordCnt = this.titleBar.recordCount.toString();
        mergeSetting.RecordId = '-1';
    }
    mergeSetting.DataSource = this.jsonDataSource;
    jsonObject.MergeSettings.push(mergeSetting);
    var dataForMerge = JSON.stringify(jsonObject);
    var fileReader = new FileReader();
    fileReader.onload = function () {
        var base64String = fileReader.result;
        var data = {
            fileName:  obj.container.documentEditor.documentName + '.docx',
            documentData: base64String,
            mergeData: dataForMerge,
            recordCount: recordCnt
        }
        obj.titleBar.showHideWaitingIndicator(true)
        var httpRequest = new XMLHttpRequest();
        httpRequest.open('Post', obj.serviceLink+'MailMergeReport', true);
        httpRequest.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
        httpRequest.onreadystatechange = function () {
            if (httpRequest.readyState === 4) {
                if (httpRequest.status === 200 || httpRequest.status === 304) {
                    if (obj.titleBar.isMergePreview) {
                        sfdtFiles = JSON.parse(httpRequest.responseText);
                        sessionStorage.setItem("sfdtFiles", JSON.stringify(sfdtFiles));                                
                        obj.previewCurrentRecord();
                    }
                    else {
                      obj.titleBar.isOpenInternal = true;
                      obj.container.documentEditor.open(httpRequest.responseText);
                        obj.titleBar.isOpenInternal = false;
                        //Hides the list of fields, while displaying the final report.
                        obj.titleBar.hideFieldList();
                    }
                } else {
                    // Failed to merge document
                    DialogUtility.alert({
                        title: 'Information',
                        content: 'Failed to merge document.',
                        showCloseIcon: true,
                        closeOnEscape: true,
                    });
                }
                obj.titleBar.showHideWaitingIndicator(false);
            }
        };
        httpRequest.send(JSON.stringify(data));
    };
    fileReader.readAsDataURL(blob);
}


showInsertFielddialog(): void {
    let instance: any = this;
    if (document.getElementById('insert_merge_field') === null || document.getElementById('insert_merge_field') === undefined) {
        let fieldcontainer: any = document.createElement('div');
        fieldcontainer.id = 'insert_merge_field';
        document.body.appendChild(fieldcontainer);
        this.insertFieldDialogObj.appendTo('#insert_merge_field');
        fieldcontainer.parentElement.style.position = 'fixed';
        fieldcontainer.style.width = 'auto';
        fieldcontainer.style.height = 'auto';
    }
    this.insertFieldDialogObj.close = (): void => { this.container.documentEditor.focusIn(); };
    this.insertFieldDialogObj.beforeOpen = (): void => { this.container.documentEditor.focusIn(); };
    this.insertFieldDialogObj.show();
    let fieldNameTextBox: any = document.getElementById('field_text');
    fieldNameTextBox.value = '';
}
closeFieldDialog(): void {
    this.insertFieldDialogObj.hide();
    this.container.documentEditor.focusIn();
}
insertFieldDialogObj: Dialog = new Dialog({
    header: 'Merge Field',
    content:
        '<div class="dialogContent">'
        // tslint:disable-next-line:max-line-length
        + '<label class="e-insert-field-label">Name:</label></br><input type="text" id="field_text" class="e-input" placeholder="Type a field to insert eg. FirstName">'
        + '</div>',
    showCloseIcon: true,
    isModal: true,
    width: 'auto',
    height: 'auto',
    close: this.closeFieldDialog,
    buttons: [
        {
            'click': (): void => {
                let fieldNameTextBox: any = document.getElementById('field_text');
                let fieldName: any = fieldNameTextBox.value;
                if (fieldName !== '') {
                    this.container.documentEditor.editor.insertField('MERGEFIELD ' + fieldName + ' \\* MERGEFORMAT');
                }
                this.insertFieldDialogObj.hide();
                this.container.documentEditor.focusIn();
            },
            buttonModel: {
                content: 'Ok',
                cssClass: 'e-flat',
                isPrimary: true,
            },
        },
        {
            'click': (): void => {
                this.insertFieldDialogObj.hide();
                this.container.documentEditor.focusIn();
            },
            buttonModel: {
                content: 'Cancel',
                cssClass: 'e-flat',
            },
        },
    ],
});
  openTemplate(): void {
    var uploadDocument = new FormData();
    uploadDocument.append('DocumentName', 'Template_Letter.doc');
    var loadDocumentUrl = this.serviceLink + 'LoadDocument';
    var httpRequest = new XMLHttpRequest();
    httpRequest.open('POST', loadDocumentUrl, true);
    var dataContext = this;
    httpRequest.onreadystatechange = function () {
      if (httpRequest.readyState === 4) {
        if (httpRequest.status === 200 || httpRequest.status === 304) {
          //Opens the SFDT for the specified file received from the web API.
          dataContext.container.documentEditor.open(httpRequest.responseText);
        }
      }
    };
    //Sends the request with template file name to web API. 
    httpRequest.send(uploadDocument);
  }  
  public fields:object = { tooltip:"category"};
  onSelect(args: any): void {
      let fieldName: any = args.item.textContent;
      this.insertField(fieldName);
  }
   insertField(fieldName: any): void {
      let fileName: any = fieldName.replace(/\n/g, '').replace(/\r/g, '').replace(/\r\n/g, '');
      let fieldCode: any = 'MERGEFIELD  ' + fileName + '  \\* MERGEFORMAT ';
      this.container.documentEditor.editor.insertField(fieldCode, '«' + fieldName + '»');
      this.container.documentEditor.focusIn();
  }
  onLoadJSON(args) {
    var obj= this;
    if (args.files[0]) {
        var file = args.files[0];
        if (file) {
            if (file.name.substr(file.name.lastIndexOf('.')) === '.json') {
                var fileReader_1 = new FileReader();
                fileReader_1.onload = function () {
                    var fileText = fileReader_1.result;
                    if (fileText) {
                      obj.loadJSON(fileText);
                    }
                };
                fileReader_1.readAsText(file);
            }
        }
    }
}
parseChildItems(childItems, tempListData, listData) {
  var obj=this;
    if (Array.isArray(childItems)) {
        if (obj.titleBar.recordCount == -1) {
          obj.titleBar.recordCount = childItems.length;
        }
        for (let i = 0; i < childItems.length; i++) {
          this.parseObject(childItems[i], tempListData, listData);
        }
    }
    else {
      this.parseObject(childItems, tempListData, listData);
    }
}
parseObject(jsonObject, tempListData, listData) {
  var obj=this;
    Object.keys(jsonObject).forEach(function (key) {
        var valueType = typeof jsonObject[key];
        var keyToAdd = key;
        if (valueType === "object") {
            keyToAdd = "TableStart:" + key;
            if (tempListData.length == 0) {
              obj.titleBar.recordCount = -1;
            }
        }
        if (tempListData.indexOf(keyToAdd) == -1) {
            tempListData.push(keyToAdd);
            listData.push({
                text: keyToAdd,
                category: 'Drag or click the field to insert.',
                htmlAttributes: { draggable: true }
            });
        }
        if (valueType === "object") {
          obj.parseChildItems(jsonObject[key], tempListData, listData);
            keyToAdd = "TableEnd:" + key;
            if (tempListData.indexOf(keyToAdd) == -1) {
                tempListData.push(keyToAdd);
                listData.push({
                    text: keyToAdd,
                    category: 'Drag or click the field to insert.',
                    htmlAttributes: { draggable: true }
                });
            }
        }
    });
}
loadJSON(data) {
    let jsonData = JSON.parse(data);
    let listData = [];
    this.tempListData = [];
    this.titleBar.recordCount = 1;
    this.jsonDataSource = jsonData;
    this.parseChildItems(jsonData, this.tempListData, listData);
    var listview = new ListView({
        //set the data to datasource property
        dataSource: listData,
        fields: { tooltip: 'category' },
        select: onselect
    });
    //Clears the old content from the div, before updating with new field names.
    (document.getElementById('listview') as any).replaceChildren();
    //Render initialized ListView
    listview.appendTo("#listview");
}
}