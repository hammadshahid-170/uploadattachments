import * as React from 'react';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { IUploadDocumentsDataProvider } from '../sharePointDataProvider/IUploadDocumentsDataProvider';
import { IUploadDocuments } from '../Models/IUploadDocuments';
import { UploadDocumentsDataProvider } from '../sharePointDataProvider/UploadDocumentsDataProvider';
import { PanelType, Panel,DefaultButton } from 'office-ui-fabric-react';

import FormUploadDocumentsCreate from '../create/FormUploadDocumentsCreate';
import FormUploadDocumentsEdit from '../edit/FormUploadDocumentsEdit';
export interface ICommandBarUploadDocumentsState {
  isVisible: boolean;
  docsupload: IUploadDocuments;
  messageSended: boolean;
  uploadDocumentsDataProvider:IUploadDocumentsDataProvider;
  _goBack:VoidFunction;
  _reload:VoidFunction;
}
export class CommandBarUploadDocuments extends React.Component<{}, ICommandBarUploadDocumentsState> {

  private  _uploadDocumentsDataProvider:IUploadDocumentsDataProvider;
  private  _docsupload:IUploadDocuments;
  /**
   *Cosnstructor og CommandBarCustomers
   */
  constructor(props) {
    super(props);
    this._uploadDocumentsDataProvider=new UploadDocumentsDataProvider({});
    this.state = {
      isVisible: false,
      docsupload: this._docsupload,
      uploadDocumentsDataProvider: this._uploadDocumentsDataProvider,
      messageSended: false,
      _goBack:this._hidePanel,
      _reload:props.state._goBack,
    };
  }

  public render(): JSX.Element {
    return (
      <div className="d-inline-block my-3">

        <DefaultButton
          className={`customStyle`}
          iconProps={{ iconName: "Add" }}
          style={{ background: 'none', borderColor: '#ff582b', color: '#ff582b', borderRadius: '5px',cursor:"pointer" }}
          onClick={() => {
            this.setState({ isVisible: true });
          }}
        >Upload Document</DefaultButton>
        <Panel isOpen={this.state.isVisible} onDismiss={this._hidePanel} type={PanelType.medium} headerText={"Upload Document"}>
        <FormUploadDocumentsCreate {...this}  />
        </Panel>
      </div>
    );
  }

  // Data for CommandBar
  private getItems = () => {
    return [
      {
        key: 'newItem',
        name: 'New',
        cacheKey: 'myCacheKey', // changing this key will invalidate this items cache
        iconProps: {
          iconName: 'Add'
        },
        ariaLabel: 'New',
        subMenuProps: {
          items: [
            {
              key: 'customerItem',
              name: 'Customer Item',
              iconProps: {
                iconName: 'SwayLogo16'
              },
              ['data-automation-id']: 'newEmailButton',
              onClick: () => {
                
                  this.setState( {isVisible:true});
                    
              }
            }
          ]
        }
      },
      {
        key: 'share',
        name: 'Share',
        iconProps: {
          iconName: 'Share'
        },
        onClick: () => console.log('Share')
      },
      {
        key: 'download',
        name: 'Export to Excel',
        iconProps: {
          iconName: 'ExcelLogo'
        },
        onClick: () => console.log('Download')
      }
    ];
  }

  private getOverlflowItems = () => {
    return [
      {
        key: 'move',
        name: 'Move to...',
        onClick: () => console.log('Move to'),
        iconProps: {
          iconName: 'MoveToFolder'
        }
      },
      {
        key: 'copy',
        name: 'Copy to...',
        onClick: () => console.log('Copy to'),
        iconProps: {
          iconName: 'Copy'
        }
      },
      {
        key: 'rename',
        name: 'Rename...',
        onClick: () => console.log('Rename'),
        iconProps: {
          iconName: 'Edit'
        }
      }
    ];
  }

  private getFarItems = () => {
    return [
      {
        key: 'sort',
        name: 'Sort',
        ariaLabel: 'Sort',
        iconProps: {
          iconName: 'SortLines'
        },
        onClick: () => console.log('Sort')
      },
      {
        key: 'tile',
        name: 'Grid view',
        ariaLabel: 'Grid view',
        iconProps: {
          iconName: 'Tiles'
        },
        iconOnly: true,
        onClick: () => console.log('Tiles')
      },
      {
        key: 'info',
        name: 'Info',
        ariaLabel: 'Info',
        iconProps: {
          iconName: 'Info'
        },
        iconOnly: true,
        onClick: () => console.log('Info')
      }
    ];
  }
  private _hidePanel = () => {
    this.setState({ isVisible: false });
  }


}