import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";


export default class SPHelper {


    public demoFunction = async () => {
        let currentUser = await this.getCurrentUserInfo();
        console.log(currentUser);
        let azUserInfo = await graph.users.getById('revathy@o365practice.onmicrosoft.com').select('employeeId', 'displayName').get();
        console.log(azUserInfo);
    }

    public getCurrentUserInfo = async () => {
        let currentUserInfo = await sp.web.currentUser.get();
        return ({
            ID: currentUserInfo.Id,
            Email: currentUserInfo.Email,
            LoginName: currentUserInfo.LoginName,
            DisplayName: currentUserInfo.Title,
            Picture: '/_layouts/15/userphoto.aspx?size=S&username=' + currentUserInfo.UserPrincipalName,
        });
    }
}