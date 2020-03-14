import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import { ISpupsProperySyncProps } from './ISpupsProperySyncProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import SPHelper from '../../../Common/SPHelper';

export default class SpupsProperySync extends React.Component<ISpupsProperySyncProps, {}> {
  private helper: SPHelper = new SPHelper();
  constructor(props: ISpupsProperySyncProps) {
    super(props);
  }

  public componentDidMount = async () => {
    this.helper.demoFunction();
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
