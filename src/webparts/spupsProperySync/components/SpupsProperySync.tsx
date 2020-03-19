import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import { ISpupsProperySyncProps } from './ISpupsProperySyncProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import SPHelper from '../../../Common/SPHelper';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { IPropertyMappings } from '../../../Common/IModel';
import PropertyMappingList from './PropertyMappingList';

export interface ISpupsProperySyncState {
  propertyMappings: IPropertyMappings[];
  uploadedTemplate?: IFilePickerResult;
}

export default class SpupsProperySync extends React.Component<ISpupsProperySyncProps, ISpupsProperySyncState> {
  private helper: SPHelper = null;
  constructor(props: ISpupsProperySyncProps) {
    super(props);
    this.state = {
      propertyMappings: [],
    };
    this.helper = new SPHelper(this.props.context.pageContext.legacyPageContext.siteAbsoluteUrl,
      this.props.context.pageContext.legacyPageContext.tenantDisplayName,
      this.props.context.pageContext.legacyPageContext.webDomain,
      this.props.context.pageContext.web.serverRelativeUrl
    );
  }

  public componentDidMount = async () => {
    let propertyMappings: IPropertyMappings[] = await this.helper.getPropertyMappings();    
    this.setState({ propertyMappings });
  }

  private uploadDataToSync = async () => {
    const { uploadedTemplate } = this.state;
    let dataToSync = await uploadedTemplate.downloadFileContent();
    let reader = new FileReader();
    reader.readAsText(dataToSync);
    reader.onload = async () => {
      console.log(reader.result);
    }
  }

  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
  }

  public render(): React.ReactElement<ISpupsProperySyncProps> {
    const { propertyMappings } = this.state;
    return (
      <div className={styles.spupsProperySync}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <PropertyMappingList mappingProperties={propertyMappings} helper={this.helper} />              
              <FilePicker
                bingAPIKey="<BING API KEY>"
                accepts={[".json"]}
                buttonIcon="FileImage"
                onSave={(uploadedTemplate: IFilePickerResult) => { this.setState({ uploadedTemplate }) }}
                onChanged={(uploadedTemplate: IFilePickerResult) => { this.setState({ uploadedTemplate }) }}
                context={this.props.context}
              />
              {this.state.uploadedTemplate &&
                <div style={{ color: "black" }}>
                  <FileTypeIcon type={IconType.font} path={this.state.uploadedTemplate.fileAbsoluteUrl ? this.state.uploadedTemplate.fileAbsoluteUrl : this.state.uploadedTemplate.fileName} />
                  &nbsp;{this.state.uploadedTemplate.fileName}
                </div>
              }
              <PrimaryButton text="Upload Data to Sync" onClick={this.uploadDataToSync} />
              <PeoplePicker
                context={this.props.context}
                titleText="People Picker"
                personSelectionLimit={20}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                isRequired={true}
                selectedItems={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
