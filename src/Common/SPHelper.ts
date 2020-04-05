import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
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
import { IUserInfo, IPropertyMappings, IPropertyPair, FileContentType } from "./IModel";
import * as _ from 'lodash';


export interface ISPHelper {
    demoFunction: () => void;
    getCurrentUserInfo: () => Promise<IUserInfo>;
    getAzurePropertyForUsers: (selectFields: string, filterQuery: string) => Promise<any[]>;
    getPropertyMappings: () => Promise<any[]>;
    getPropertyMappingsTemplate: (propertyMappings: IPropertyMappings[]) => Promise<any>;
    addFilesToFolder: (filename: string, fileContent: any) => void;
    getFileContent: (filepath: string, contentType: FileContentType) => void;

    runAzFunction: (httpClient: HttpClient, inputData: any) => void;
}

export default class SPHelper implements ISPHelper {

    private SiteURL: string = "";
    private SiteRelativeURL: string = "";
    private AdminSiteURL: string = "";
    private SyncTemplateFilePath: string = "/Shared Documents/SyncJobTemplate/";
    private SyncUploadFilePath: string = "/Shared Documents/UPSDataToProcess/";
    private SyncJSONFileName: string = `SyncTemplate_${moment().format("MM-DD-YYYY-HH-mm-ss")}.json`;
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
        let azUserInfo = await graph.users
            .filter(`userPrincipalName eq 'AdeleV@o365practice.onmicrosoft.com' or userPrincipalName eq 'AlexW@o365practice.onmicrosoft.com'`)
            .select('employeeId', 'displayName', 'city', 'state').get();
        console.log(azUserInfo);
    }
    /**
     * Get the Azure property data for the Users
     */
    public getAzurePropertyForUsers = async (selectFields: string, filterQuery: string): Promise<any[]> => {
        let users = await graph.users.filter(filterQuery).select(selectFields).get();
        return _.orderBy(users, 'displayName', 'asc');
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
        let propertyPair: any[] = [];
        let sampleUser1 = new Object();
        let sampleUser2 = new Object();
        sampleUser1['UserID'] = "user1@tenantname.onmicrosoft.com";
        sampleUser2['UserID'] = "user2@tenantname.onmicrosoft.com";
        propertyMappings.map((propsMap: IPropertyMappings) => {
            sampleUser1[propsMap.SPProperty] = "";
            sampleUser2[propsMap.SPProperty] = "";
        });
        propertyPair.push(sampleUser1, sampleUser2);
        finalJson = JSON.stringify(propertyPair);
        return JSON.parse(finalJson);
    }
    public getPropertyMappingsTemplate1 = async (propertyMappings: IPropertyMappings[]) => {
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
    public getFileContent = async (filepath: string, contentType: FileContentType) => {
        switch (contentType) {
            case FileContentType.Blob:
                return await this._web.getFileByServerRelativeUrl(filepath).getBlob();
            case FileContentType.ArrayBuffer:
                return await this._web.getFileByServerRelativeUrl(filepath).getBuffer();
            case FileContentType.Text:
                return await this._web.getFileByServerRelativeUrl(filepath).getText();
            case FileContentType.JSON:
                return await this._web.getFileByServerRelativeUrl(filepath).getJSON();
        }
    }
    /**
     * Add the template file to a folder with contents.
     * This is used for creating the template json file.
     */
    public addFilesToFolder = async (fileContent: any, isCSV: boolean) => {
        let filename = (isCSV) ? this.SyncCSVFileName : this.SyncJSONFileName;
        await this.checkAndCreateFolder(this.SiteRelativeURL + this.SyncTemplateFilePath);
        return await this._web.getFolderByServerRelativeUrl(this.SiteRelativeURL + this.SyncTemplateFilePath)
            .files
            .add(decodeURI(this.SiteRelativeURL + this.SyncTemplateFilePath + filename), fileContent, true);
    }
    /**
     * Add the data file to a folder with contents.
     * This is used for creating the template json file.
     */
    public addDataFilesToFolder = async (fileContent: any, filename: string) => {
        await this.checkAndCreateFolder(this.SiteRelativeURL + this.SyncUploadFilePath);
        return await this._web.getFolderByServerRelativeUrl(this.SiteRelativeURL + this.SyncUploadFilePath)
            .files
            .add(decodeURI(this.SiteRelativeURL + this.SyncUploadFilePath + filename), fileContent, true);
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

    protected functionUrl: string = "https://demosponline.azurewebsites.net/api/playwithpnpsp?code=mdEonK9e7eS38WziRbdllF19StdOFQQhquAbhSUivMbX8vgjQ1GNPg==";
    public runAzFunction = async (httpClient: HttpClient, inputData: any) => {
        const requestHeaders: Headers = new Headers();
        requestHeaders.append("Content-type", "application/json");
        requestHeaders.append("Cache-Control", "no-cache");
        const postOptions: IHttpClientOptions = {
            headers: requestHeaders,
            body: `${inputData}`
        };
        let response: HttpClientResponse = await httpClient.post(this.functionUrl, HttpClient.configurations.v1, postOptions);
        console.log("Actual Response: ", response);
    }

}