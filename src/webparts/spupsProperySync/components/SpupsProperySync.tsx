import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import { ISpupsProperySyncProps } from './ISpupsProperySyncProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import SPHelper, { IPropertyMappings } from '../../../Common/SPHelper';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";

export interface ISpupsProperySyncState {
  showDownloadLink: boolean;
  downloadLink: string;
  templateFileName: string;
  uploadedTemplate?: IFilePickerResult;
}

export default class SpupsProperySync extends React.Component<ISpupsProperySyncProps, ISpupsProperySyncState> {
  private helper: SPHelper = null;
  constructor(props: ISpupsProperySyncProps) {
    super(props);
    this.state = {
      showDownloadLink: false,
      downloadLink: "",
      templateFileName: ""
    };
    this.helper = new SPHelper(this.props.context.pageContext.legacyPageContext.siteAbsoluteUrl,
      this.props.context.pageContext.legacyPageContext.tenantDisplayName,
      this.props.context.pageContext.legacyPageContext.webDomain,
      this.props.context.pageContext.web.serverRelativeUrl
    );
  }

  public componentDidMount = async () => {

  }

  private generatePropertyMappingTemplate = async () => {
    let jsonOut = await this.helper.getPropertyMappings();
    let fileTemplate = await this.helper.addFilesToFolder(JSON.stringify(jsonOut));
    this.setState({
      downloadLink: fileTemplate.data.ServerRelativeUrl,
      templateFileName: fileTemplate.data.Name,
      showDownloadLink: true
    });
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

  private getJSONFile = async () => {
    let blobContent: any = await this.helper.getFileContentAsBlob(this.state.downloadLink);
    if (window.navigator.msSaveOrOpenBlob) {
      window.navigator.msSaveBlob(blobContent, this.state.templateFileName);
    } else {
      const anchor = window.document.createElement('a');
      anchor.href = window.URL.createObjectURL(blobContent);
      anchor.download = this.state.templateFileName;
      document.body.appendChild(anchor);
      anchor.click();
      document.body.removeChild(anchor);
    }
  }

  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
  }

  public render(): React.ReactElement<ISpupsProperySyncProps> {
    return (
      <div className={styles.spupsProperySync}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <PrimaryButton text="Generate Template" onClick={this.generatePropertyMappingTemplate} />
              {this.state.showDownloadLink &&
                <a href={"javascript:void(0);"} onClick={this.getJSONFile}>{"Download Template"}</a>
              }
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
