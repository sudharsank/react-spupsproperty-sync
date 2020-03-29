import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import * as strings from 'SpupsProperySyncWebPartStrings';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { FileTypeIcon, IconType } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { IPropertyMappings, FileContentType } from '../../../Common/IModel';
//import { ISpupsProperySyncProps } from './ISpupsProperySyncProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import SPHelper from '../../../Common/SPHelper';
import PropertyMappingList from './PropertyMapping/PropertyMappingList';
import UPPropertyData from './UPPropertyData';
import ManualPropertyUpdate from './ManualPropertyUpdate';
import AzurePropertyView from './AzurePropertyView';
import * as _ from 'lodash';

export interface ISpupsProperySyncProps {
    context: WebPartContext;
}

export interface ISpupsProperySyncState {
    propertyMappings: IPropertyMappings[];
    uploadedTemplate?: IFilePickerResult;
    uploadedFileURL?: string;
    showUploadData: boolean;
    showUploadProgress: boolean;
    showPropsLoader: boolean;
    disablePropsButtons: boolean;
    uploadedData?: any;
    isCSV: boolean;
    selectedUsers?: any[];
    manualPropertyData: any[];
    azurePropertyData: any[];
}

export default class SpupsProperySync extends React.Component<ISpupsProperySyncProps, ISpupsProperySyncState> {
    // Private variables
    private helper: SPHelper = null;
    /**
     * Constructor
     * @param props 
     */
    constructor(props: ISpupsProperySyncProps) {
        super(props);
        this.state = {
            propertyMappings: [],
            showUploadData: false,
            showUploadProgress: false,
            showPropsLoader: false,
            disablePropsButtons: false,
            isCSV: false,
            selectedUsers: [],
            manualPropertyData: [],
            azurePropertyData: []
        };
        this.helper = new SPHelper(this.props.context.pageContext.legacyPageContext.siteAbsoluteUrl,
            this.props.context.pageContext.legacyPageContext.tenantDisplayName,
            this.props.context.pageContext.legacyPageContext.webDomain,
            this.props.context.pageContext.web.serverRelativeUrl
        );
    }
    /**
     * Component mount
     */
    public componentDidMount = async () => {
        let propertyMappings: IPropertyMappings[] = await this.helper.getPropertyMappings();
        propertyMappings.map(prop => { prop.IsIncluded = true; });
        this.setState({ propertyMappings });
        //this.helper.demoFunction();
    }
    /**
     * Uploading data file and displaying the contents of the file
     */
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
                this.setState({ showUploadProgress: false, uploadedData: filecontent, isCSV: ext.toLocaleLowerCase() == "csv" });
            } else {
                let dataToSync = await uploadedTemplate.downloadFileContent();
                let filereader = new FileReader();
                filereader.readAsBinaryString(dataToSync);
                filereader.onload = async () => {
                    let dataUploaded = await this.helper.addDataFilesToFolder(filereader.result, uploadedTemplate.fileName);
                    if (ext.toLocaleLowerCase() == "csv") {
                        filecontent = await dataUploaded.file.getText();
                    } else if (ext.toLocaleLowerCase() == "json") {
                        filecontent = await dataUploaded.file.getJSON();
                    }
                    this.setState({ showUploadProgress: false, uploadedData: filecontent, isCSV: ext.toLocaleLowerCase() == "csv" });
                };
            }
        }
    }
    /**
     * Triggers when the users are selected for manual update
     */
    private _getPeoplePickerItems = (items: any[]) => {
        this.setState({ selectedUsers: items });
    }
    /**
     * Display the inline editing table to edit the properties for manual update
     */
    private _getManualPropertyTable = () => {
        this.setState({ disablePropsButtons: true, showPropsLoader: true });
        const { propertyMappings, selectedUsers } = this.state;
        let includedProperties: IPropertyMappings[] = propertyMappings.filter((o) => { return o.IsIncluded; });
        let manualPropertyData: any[] = [];
        if (selectedUsers && selectedUsers.length > 0) {
            selectedUsers.map(user => {
                let userObj = new Object();
                userObj['UserID'] = user.loginName;
                userObj['DisplayName'] = user.text;
                userObj['ImageUrl'] = user.imageUrl;
                includedProperties.map((propsMap: IPropertyMappings) => {
                    userObj[propsMap.SPProperty] = "";
                });
                manualPropertyData.push(userObj);
            });
            this.setState({ manualPropertyData, azurePropertyData: [], showPropsLoader: false, disablePropsButtons: false });
        } else {
            this.setState({ disablePropsButtons: false, showPropsLoader: false });
        }
    }
    /**
     * Get the property values from Azure
     */
    private _getAzurePropertyTable = async () => {
        this.setState({ disablePropsButtons: true, showPropsLoader: true });
        const { propertyMappings, selectedUsers } = this.state;
        let includedProperties: IPropertyMappings[] = propertyMappings.filter((o) => { return o.IsIncluded; });
        let selectFields: string = "id, userPrincipalName, displayName, " + _.map(includedProperties, 'AzProperty').join(',');
        let tempQuery: string[] = []; let filterQuery: string = ``;
        if (selectedUsers && selectedUsers.length > 0) {
            selectedUsers.map(user => {
                tempQuery.push(`userPrincipalName eq '${user.loginName.split('|')[2]}'`);
            });
            filterQuery = tempQuery.join(' or ');
            let azurePropertyData = await this.helper.getAzurePropertyForUsers(selectFields, filterQuery);
            this.setState({ azurePropertyData, manualPropertyData: [], showPropsLoader: false, disablePropsButtons: false });
        } else {
            this.setState({ disablePropsButtons: false, showPropsLoader: false });
        }
    }
    /**
     * On selecting the data file for update
     */
    private _onSaveTemplate = (uploadedTemplate: IFilePickerResult) => {
        this.setState({ uploadedTemplate, showUploadData: true });
    }
    /**
     * On changing the data file for update
     */
    private _onChangeTemplate = (uploadedTemplate: IFilePickerResult) => {
        this.setState({ uploadedTemplate, showUploadData: true });
    }
    /**
     * Update with manual properties
     */
    private _updateSPWithManualProperties = async (data: any[]) => {
        console.log("_updateSPWithManualProperties: ", data);
    }
    /**
     * Update with azure properties
     */
    private _updateSPWithAzureProperties = async (data: any[]) => {
        console.log("_updateSPWithAzureProperties: ", data);
    }
    /**
     * Component render
     */
    public render(): React.ReactElement<ISpupsProperySyncProps> {
        const { propertyMappings, uploadedTemplate, uploadedFileURL, showUploadData, showUploadProgress, uploadedData, isCSV, selectedUsers, manualPropertyData,
            azurePropertyData, disablePropsButtons, showPropsLoader } = this.state;
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
                                titleText={strings.PPLPickerTitleText}
                                personSelectionLimit={20}
                                groupName={""} // Leave this blank in case you want to filter from all users
                                showtooltip={false}
                                isRequired={false}
                                selectedItems={this._getPeoplePickerItems}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={500} />
                            {selectedUsers && selectedUsers.length > 0 &&
                                <div style={{ marginTop: "5px" }}>
                                    <PrimaryButton text={strings.BtnManualProps} onClick={this._getManualPropertyTable} style={{ marginRight: '5px' }} disabled={disablePropsButtons} />
                                    <PrimaryButton text={strings.BtnAzureProps} onClick={this._getAzurePropertyTable} disabled={disablePropsButtons} />
                                    {showPropsLoader && <Spinner className={styles.generateTemplateLoader} label={strings.PropsLoader} ariaLive="assertive" labelPosition="right" />}
                                </div>
                            }
                            {manualPropertyData && manualPropertyData.length > 0 &&
                                <ManualPropertyUpdate userProperties={manualPropertyData} UpdateSPUserWithManualProps={this._updateSPWithManualProperties} />
                            }
                            {azurePropertyData && azurePropertyData.length > 0 &&
                                <AzurePropertyView userProperties={azurePropertyData} UpdateSPUserWithAzureProps={this._updateSPWithAzureProperties} />
                            }
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}
