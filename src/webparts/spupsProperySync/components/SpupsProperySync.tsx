import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import * as strings from 'SpupsProperySyncWebPartStrings';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { IPropertyMappings, FileContentType } from '../../../Common/IModel';
import { ISpupsProperySyncProps } from './ISpupsProperySyncProps';
import SPHelper from '../../../Common/SPHelper';
import PropertyMappingList from './PropertyMappingList';
import UPPropertyData from './UPPropertyData';

export interface ISpupsProperySyncState {
    propertyMappings: IPropertyMappings[];
    uploadedTemplate?: IFilePickerResult;
    uploadedFileURL?: string;
    showUploadData: boolean;
    showUploadProgress: boolean;
    uploadedData?: any;
    isCSV: boolean;
}

export default class SpupsProperySync extends React.Component<ISpupsProperySyncProps, ISpupsProperySyncState> {
    private helper: SPHelper = null;
    constructor(props: ISpupsProperySyncProps) {
        super(props);
        this.state = {
            propertyMappings: [],
            showUploadData: false,
            showUploadProgress: false,
            isCSV: false
        };
        this.helper = new SPHelper(this.props.context.pageContext.legacyPageContext.siteAbsoluteUrl,
            this.props.context.pageContext.legacyPageContext.tenantDisplayName,
            this.props.context.pageContext.legacyPageContext.webDomain,
            this.props.context.pageContext.web.serverRelativeUrl
        );
    }

    public componentDidMount = async () => {
        //console.log(this.props.context.pageContext.legacyPageContext);
        let propertyMappings: IPropertyMappings[] = await this.helper.getPropertyMappings();
        this.setState({ propertyMappings });
    }

    private uploadDataToSync = async () => {
        this.setState({ showUploadProgress: true });
        const { uploadedTemplate } = this.state;
        let filecontent: any = null;
        if (uploadedTemplate && uploadedTemplate.fileName) {
            let ext: string = uploadedTemplate.fileName.split('.').pop();
            if (uploadedTemplate.fileAbsoluteUrl && null !== uploadedTemplate.fileAbsoluteUrl) {
                let filerelativeurl: string = "";
                if (uploadedTemplate.fileAbsoluteUrl.indexOf(this.props.context.pageContext.legacyPageContext.webAbsoluteUrl) >= 0) {
                    filerelativeurl = uploadedTemplate.fileAbsoluteUrl.replace(this.props.context.pageContext.legacyPageContext.webAbsoluteUrl,
                        this.props.context.pageContext.legacyPageContext.webServerRelativeUrl);
                }
                filecontent = await this.helper.getFileContent(filerelativeurl, FileContentType.Blob);
                await this.helper.addDataFilesToFolder(filecontent, uploadedTemplate.fileName);
                if (ext.toLocaleLowerCase() == "csv") {
                    filecontent = await this.helper.getFileContent(filerelativeurl, FileContentType.Text);
                } else if (ext.toLocaleLowerCase() == "json") {
                    filecontent = await this.helper.getFileContent(filerelativeurl, FileContentType.JSON);
                }
                console.log(filecontent);
                this.setState({ showUploadProgress: false, uploadedData: filecontent, isCSV: ext.toLocaleLowerCase() == "csv" });
            } else {
                let dataToSync = await uploadedTemplate.downloadFileContent();
                let filereader = new FileReader();
                filereader.readAsBinaryString(dataToSync);
                filereader.onload = async () => {
                    let dataUploaded = await this.helper.addDataFilesToFolder(filereader.result, uploadedTemplate.fileName);
                    //this.setState({ uploadedFileURL: dataUploaded.data.ServerRelativeUrl });
                    if (ext.toLocaleLowerCase() == "csv") {
                        filecontent = await dataUploaded.file.getText();
                    } else if (ext.toLocaleLowerCase() == "json") {
                        filecontent = await dataUploaded.file.getJSON();
                    }
                    console.log(filecontent);
                    this.setState({ showUploadProgress: false, uploadedData: filecontent, isCSV: ext.toLocaleLowerCase() == "csv" });
                };
            }
        }
    }

    private _getPeoplePickerItems(items: any[]) {
        console.log('Items:', items);
    }

    private _onSaveTemplate = (uploadedTemplate: IFilePickerResult) => {
        this.setState({ uploadedTemplate, showUploadData: true });
    }

    private _onChangeTemplate = (uploadedTemplate: IFilePickerResult) => {
        this.setState({ uploadedTemplate, showUploadData: true });
    }

    public render(): React.ReactElement<ISpupsProperySyncProps> {
        const { propertyMappings, uploadedTemplate, uploadedFileURL, showUploadData, showUploadProgress, uploadedData, isCSV } = this.state;
        const fileurl = uploadedFileURL ? uploadedFileURL : uploadedTemplate && uploadedTemplate.fileAbsoluteUrl ? uploadedTemplate.fileAbsoluteUrl :
            uploadedTemplate && uploadedTemplate.fileName ? uploadedTemplate.fileName : '';
        return (
            <div className={styles.spupsProperySync}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <PropertyMappingList mappingProperties={propertyMappings} helper={this.helper} disabled={showUploadProgress} />
                            <div>
                                <FilePicker
                                    accepts={[".json", ".csv"]}
                                    buttonIcon="FileImage"
                                    onSave={this._onSaveTemplate}
                                    onChanged={this._onChangeTemplate}
                                    context={this.props.context}
                                    disabled={showUploadProgress}
                                    buttonLabel={"Select Data file"}
                                    hideLinkUploadTab={true}
                                    hideOrganisationalAssetTab={true}
                                    hideWebSearchTab={true}
                                />
                            </div>
                            {fileurl &&
                                <div style={{ color: "black" }}>
                                    <FileTypeIcon type={IconType.font} path={fileurl} />
                                    &nbsp;{uploadedTemplate.fileName}
                                </div>
                            }
                            {showUploadData &&
                                <PrimaryButton text={strings.BtnUploadDataForSync} onClick={this.uploadDataToSync} disabled={showUploadProgress} />
                            }
                            {showUploadProgress && <Spinner className={styles.generateTemplateLoader} label={strings.UploadDataToSyncLoader} ariaLive="assertive" labelPosition="right" />}
                            <UPPropertyData items={uploadedData} isCSV={isCSV} />
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
