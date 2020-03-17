import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import { ISpupsProperySyncProps } from './ISpupsProperySyncProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import SPHelper, { IPropertyMappings } from '../../../Common/SPHelper';

export default class SpupsProperySync extends React.Component<ISpupsProperySyncProps, {}> {
  private helper: SPHelper = null;
  constructor(props: ISpupsProperySyncProps) {
    super(props);
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
    console.log(fileTemplate.data.Name, fileTemplate.data.ServerRelativeUrl);
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
