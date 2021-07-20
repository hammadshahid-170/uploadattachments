import * as React from 'react';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import * as lodash from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
//import { Separator } from 'office-ui-fabric-react/lib/Separator';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
//import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { CommandBarUploadDocuments } from '../utils/CommandBarUploadDocuments';
import { IDetailsListUploadDocumentsState } from './IDetailsListUploadDocumentsState';
import { UploadDocumentsDataProvider } from '../sharePointDataProvider/UploadDocumentsDataProvider';
import { IUploadDocuments } from '../Models/IUploadDocuments';
import FormUploadDocumentsEdit from '../edit/FormUploadDocumentsEdit';
import { getId } from 'office-ui-fabric-react/lib/Utilities';

import "bootstrap/dist/css/bootstrap.min.css";
import { DefaultButton, Button } from 'office-ui-fabric-react/lib/Button';
const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px'
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden'
      }
    }
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px'
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap'
  },
  selectionDetails: {
    marginBottom: '20px'
  }
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px'
  }
};
const SITEURL = window.location.origin
export class DetailsListUploadDocuments extends React.Component<{}, IDetailsListUploadDocumentsState> {
  private _selection: Selection;
  private _allItems: IUploadDocuments[];
  private _uploadDocumentsDataProvider: UploadDocumentsDataProvider;
  private _webPartContext: IWebPartContext;
  private showEditPanel: boolean;
  // Use getId() to ensure that the IDs are unique on the page.
  // (It's also okay to use plain strings without getId() and manually ensure uniqueness.)
  private _titleId: string = getId('title');
  private _subtitleId: string = getId('subText');
  constructor(props: {}) {

    super(props);
    //this is to chage by wev service rest apiget from the list
    this._uploadDocumentsDataProvider = new UploadDocumentsDataProvider({});
    // this._allItems = 
    const columns: IColumn[] = [
      {
        key: 'column1', name: 'Delete', fieldName: 'Delete',  minWidth: 50, maxWidth: 50, isResizable: true,
        onRender: (item) => (
          <button onClick={() => {
            { if (window.confirm('Are you sure you want to delete?')) this.deleteItem(item, this); }
          }} id={item.ID} type="button" className="btn btn-danger" style={{ cursor: 'pointer', fontSize: '9px' }}>Delete</button>
        )
      },
      {
        key: 'column2', name: 'Edit', fieldName: 'Edit', minWidth: 50, maxWidth: 50, isResizable: true,
        onRender: (item) => (
          <button  onClick={() => {
               this._onItemEdit(item, this)
             }} id={item.ID} type="button" className="btn btn-primary" style={{ cursor: 'pointer', fontSize: '9px' }}>Edit</button>
        )
      },
      {
        key: 'column4',
        name: 'ID',
        isIconOnly: false,
        fieldName: 'ID',
        minWidth: 30,
        maxWidth: 50,
        data: 'string',
        onColumnClick: this._onColumnClick,
      },
      {
        key: 'column5',
        name: 'Name',
        fieldName: 'Name',
        minWidth: 100,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
       // isSorted: true,
       // isSortedDescending: false,
        // sortAscendingAriaLabel: 'Sorted A to Z',
        // sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'column6',
        name: 'Email',
        fieldName: 'Email',
        minWidth: 100,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
       // isSorted: true,
        //isSortedDescending: false,
        // sortAscendingAriaLabel: 'Sorted A to Z',
        // sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'column7',
        name: 'Type',
        fieldName: 'Type',
        minWidth: 100,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
       // isSorted: true,
        //isSortedDescending: false,
        // sortAscendingAriaLabel: 'Sorted A to Z',
        // sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'column8',
        name: 'gender',
        fieldName: 'gender',
        minWidth: 100,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
       
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      }
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
         // selectionDetails: this._getSelectionDetails(),
          showEditUploadDocumentsPanel: this.showEditPanel,

        });
      }
    });
this._allItems=[]
    this.state = {
      items: this._allItems,
      columns: columns,
      selectionDetails: this._getSelectionDetails(),
      showEditUploadDocumentsPanel: false,
      selectedUploadDocuments: null,
      _goBack: this._hidePanel,
      showModal:false
    };
  }
  componentDidMount() {
    this._LoadDocuments();
  }

  public render() {
    const { columns, items, selectionDetails, showEditUploadDocumentsPanel } = this.state;

    return (
      <Fabric>
        
        <CommandBarUploadDocuments  {...this} />
       
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={items}
            columns={columns}
            selectionMode={SelectionMode.single}
            //getKey={this._getKey}
          //  setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
          //  selection={this._selection}
        //    selectionPreservedOnEmptyClick={true}
           // onItemInvoked={(item) => { this._onItemInvoked(item, this); }}
           // enterModalSelectionOnTouch={true}
          //  ariaLabelForSelectionColumn="Toggle selection"
           // ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            //checkButtonAriaLabel="Row checkbox"
          />

        </MarqueeSelection>
        <div>
          <Panel isOpen={this.state.showEditUploadDocumentsPanel} onDismiss={this._hidePanel} type={PanelType.medium} headerText="Edit Document">
            <FormUploadDocumentsEdit {...this} />
          </Panel>
        </div>
      </Fabric>
    );
  }
  private _LoadDocuments() {
    const items: IUploadDocuments[] = [];
    this._uploadDocumentsDataProvider.getItems().then((docs: IUploadDocuments[]) => {
      docs.forEach(element => {
         items.push({
          ID: element.ID,
          Name: element.Name,
          Email: element.Email,
          Type: element.Type,
          gender: element.gender,
          Attachments: element.Attachments,
          FileNames: element.FileNames,
          ServerRelativeUrls: element.ServerRelativeUrls,
          
          });
      });
      this.setState({ items: docs })
      return docs;

    });
    return items;
  }
  private _showModal(item: IUploadDocuments) {
    let fileUrl: string;
    $("#previewAttachmentsRow").empty();
    for (let i = 0; i < item.ServerRelativeUrls.length; i++) {
      fileUrl = SITEURL + item.ServerRelativeUrls[i];
      const fileExtension = item.FileNames[i].split(".")[1].toLowerCase();
      if (fileExtension === "jpg" || fileExtension === "jpeg" || fileExtension === "png" || fileExtension === "gif" || fileExtension === "tiff") {
        $("#previewAttachmentsRow").append(
          "<div class='col-xl-6 col-lg-6 col-md-6 col-sm-12 mt-2'>" +
          "<img class='img-fluid img-thumbnail' src='" + fileUrl + "' />" +
          "<br>" +
          "<a class='d-inline-block text-truncate' style='max-width: 280px;' href='" + fileUrl + "' target='_blank' id='" + item.FileNames[i] + "'>" + item.FileNames[i] + "</a>" +
          "</div>"
        );
      }
      else {
        // if (fileExtension === "doc" || fileExtension === "docx" || fileExtension === "pdf")
        // const iconName = fileExtension === "doc" ? "Document" : fileExtension === "docx" ? "WordDocument" : fileExtension === "pdf" ? "PDF" : "";
        // "<iframe class='doc' src='" + fileUrl + "'></iframe>" +
        // <i style='font-size: x-large;' class='ms-Icon ms-Icon--" + iconName + "' aria-hidden='true'></i> 
        $("#previewAttachmentsRow").append(
          "<div class='col-xl-6 col-lg-6 col-md-6 col-sm-12 mt-2'>" +
          "<a href='" + fileUrl + "' target='_blank'>" + item.FileNames[i] + "</a>" +
          "</div>"
        );
      }
    }
    this.setState({ showModal: true });
  }
  //To Update the items in the list
  public componentDidUpdate(previousProps: any, previousState: IDetailsListUploadDocumentsState) {

  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }
  

  private _onItemInvoked(item: any, value: any): void {
    let itemDocs = item as IUploadDocuments;
    value.setState({ selectedUploadDocuments: itemDocs });
    value.setState({ showEditUploadDocumentsPanel: true });

  }


  private _hidePanel = () => {

    const items: IUploadDocuments[] = [];
    this._uploadDocumentsDataProvider.getItems().then((docs: IUploadDocuments[]) => {
      docs.forEach(element => {
         items.push({
          ID: element.ID,
          Name: element.Name,
          Email: element.Email,
          Type: element.Type,
          gender: element.gender,
          Attachments: element.Attachments,
          FileNames: element.FileNames,
          ServerRelativeUrls: element.ServerRelativeUrls,
         });
      });
      this.setState({ showEditUploadDocumentsPanel: false, items: items })
    });
  }
  private _onItemEdit(item: any, value: any): void {
    console.log("on edit clicked", item)
    let itemuploaddoc = item as IUploadDocuments;
    this._uploadDocumentsDataProvider.getItemsById(itemuploaddoc.ID).then((uploaddoc: any) => {
      let itemuploaddocs = uploaddoc[0] as IUploadDocuments;
      value.setState({ selectedUploadDocuments: itemuploaddocs });
      value.setState({ showEditUploadDocumentsPanel: true });
    });
  }
  private deleteItem(item: any, value: any): void {

    let itemuploaddoc = item as IUploadDocuments;
    this._uploadDocumentsDataProvider.deleteItem(itemuploaddoc);
    const items = this.state.items.filter(i => i.ID !== item.ID);
    this.setState({ items });
  }
  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    this.setState({ selectedUploadDocuments: this._selection.getSelection()[0] as IUploadDocuments });
    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IUploadDocuments).name;
      default:
        return `${selectionCount} items selected`;
    }
  }
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}










