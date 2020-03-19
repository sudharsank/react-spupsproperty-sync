import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists/web";
import "@pnp/sp/items/list";
import "@pnp/sp/profiles";
import "@pnp/sp/search";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { ISearchQuery, SearchResults, SearchQueryBuilder } from "@pnp/sp/search";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import * as moment from 'moment';
import { IWeb } from "@pnp/sp/webs";
import { IUserInfo, IPropertyMappings, IPropertyPair } from "./IModel";


export interface ISPHelper {
    demoFunction: () => void;
    getCurrentUserInfo: () => Promise<IUserInfo>;
    getPropertyMappings: () => Promise<any[]>;
    getPropertyMappingsTemplate: (propertyMappings: IPropertyMappings[]) => Promise<any>;
    addFilesToFolder: (filename: string, fileContent: any) => void;
    getFileContentAsBlob: (filepath: string) => void;
}

export default class SPHelper implements ISPHelper {

    private SiteURL: string = "";
    private SiteRelativeURL: string = "";
    private AdminSiteURL: string = "";
    private SyncTemplateFilePath: string = "/Shared Documents/SyncJobTemplate/";
    private SyncUploadFilePath: string = "/Shared Documents/SyncJobUploadedFiles/";
    private SyncFileName: string = `SyncTemplate_${moment().format("MM-DD-YYYY-HH-mm-ss")}.json`;
    private SyncCSVFileName: string = `SyncTemplate_${moment().format("MM-DD-YYYY-HH-mm-ss")}.csv`;
    private _web: IWeb = null;

    private Lst_PropsMapping = 'Sync Properties Mapping';
    private Lst_SyncJobs = 'UPS Sync Jobs';

    constructor(siteurl: string, tenantname: string, domainname: string, relativeurl: string) {
        this.SiteURL = siteurl;
        this.SiteRelativeURL = relativeurl;
        this.AdminSiteURL = `https://${tenantname}-admin.${domainname}`;
        this._web = sp.web;
    }

    public demoFunction = async () => {
        // let currentUser = await this.getCurrentUserInfo();
        // console.log(currentUser);
        // let azUserInfo = await graph.users.getById('revathy@o365practice.onmicrosoft.com').select('employeeId', 'displayName').get();
        // console.log(azUserInfo);
        let userToUpdate = await sp.web.siteUsers.getByEmail('revathy@o365practice.onmicrosoft.com').get();
        console.log(userToUpdate);
        // await sp.profiles.setSingleValueProfileProperty(userToUpdate.LoginName, "Title", "Revathy Sudharsan");
        // console.log("Updated");



        const result = await sp.profiles.clientPeoplePickerSearchUser({
            AllowEmailAddresses: true,
            AllowMultipleEntities: false,
            MaximumEntitySuggestions: 25,
            QueryString: 'Manager:*sudha*'
        });
        console.log(result);

        const results2: SearchResults = await sp.search("Manager:*Sudha*");
        console.log(results2);
    }
    /**
     * Get the property mappings from the 'Sync Properties Mapping' list.
     */
    public getPropertyMappings = async (): Promise<any[]> => {
        return await this._web.lists.getByTitle(this.Lst_PropsMapping).items
            .select("ID", "Title", "AzProperty", "SPProperty", "IsActive", "AutoSync")
            .filter(`IsActive eq 1`)
            .get();
    }
    /**
     * Generated the property mapping json content. 
     */
    public getPropertyMappingsTemplate = async (propertyMappings: IPropertyMappings[]) => {
        if (!propertyMappings) propertyMappings = await this.getPropertyMappings();
        let finalJson: string = "";
        let propertyPair: IPropertyPair[] = [];
        propertyMappings.map((propsMap: IPropertyMappings) => {
            propertyPair.push({
                name: propsMap.SPProperty,
                value: ""
            });
        });
        finalJson = `{
            "targetAdminUrl": "${this.AdminSiteURL}",
            "targetSiteUrl": "${this.SiteURL}",
            "values": [
                {
                    "UserID": "userid@tenantname.onmicrosoft.com",
                    "Properties": ${JSON.stringify(propertyPair)}
                }
            ]
        }`;
        return JSON.parse(finalJson);
    }
    /**
     * Get the file content as blob based on the file url.
     */
    public getFileContentAsBlob = async (filepath: string) => {
        return await this._web.getFileByServerRelativeUrl(filepath).getBlob();
    }
    /**
     * Add a file to a folder with contents.
     * This is used for creating the template json file.
     */
    public addFilesToFolder = async (fileContent: any, isCSV?: boolean) => {
        await this.checkAndCreateFolder(this.SiteRelativeURL + this.SyncTemplateFilePath);
        return await this._web.getFolderByServerRelativeUrl(this.SiteRelativeURL + this.SyncTemplateFilePath)
            .files
            .add(decodeURI(this.SiteRelativeURL + this.SyncTemplateFilePath + (isCSV) ? this.SyncCSVFileName : this.SyncFileName), fileContent, true);
    }
    /**
     * Check for the template folder, if not creates.
     */
    public checkAndCreateFolder = async (folderPath: string) => {
        try {
            await this._web.getFolderByServerRelativeUrl(folderPath).get();
        } catch (err) {
            await this._web.folders.add(folderPath);
        }
    }
    /**
     * Get current logged in user information.
     */
    public getCurrentUserInfo = async (): Promise<IUserInfo> => {
        let currentUserInfo = await this._web.currentUser.get();
        return ({
            ID: currentUserInfo.Id,
            Email: currentUserInfo.Email,
            LoginName: currentUserInfo.LoginName,
            DisplayName: currentUserInfo.Title,
            Picture: '/_layouts/15/userphoto.aspx?size=S&username=' + currentUserInfo.UserPrincipalName,
        });
    }


}