import * as React from 'react';
import * as strings from 'SpupsProperySyncWebPartStrings';
import styles from './SpupsProperySync.module.scss';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { MessageScope, IPropertyMappings } from '../../../Common/IModel';
import EditableTable from './DynamicTable/EditableTable';
import MessageContainer from './MessageContainer';

export interface IAzurePropertyViewProps {
    userProperties: any;
    showProgress: boolean;
    UpdateSPUserWithAzureProps: (data: any) => void;
    propertyMappings: IPropertyMappings[];
}

export interface IAzurePropertyViewState {
    data: any;
    propertyMappings: IPropertyMappings[];
}

export default class AzurePropertyView extends React.Component<IAzurePropertyViewProps, IAzurePropertyViewState> {
    constructor(props: IAzurePropertyViewProps) {
        super(props);
        this.state = {
            data: [],
            propertyMappings: []
        };
    }

    public componentDidMount = async () => {
        this.setState({ data: this.props.userProperties, propertyMappings: this.props.propertyMappings });
    }

    public componentDidUpdate = (prevProps: IAzurePropertyViewProps) => {
        if (prevProps.userProperties !== this.props.userProperties ||
            prevProps.propertyMappings !== this.props.propertyMappings) {
            this.setState({ data: this.props.userProperties, propertyMappings: this.props.propertyMappings });
        }
    }

    private handleRowDel = (item) => {
        var index = this.state.data.indexOf(item);
        this.state.data.splice(index, 1);
        this.setState(this.state.data);
    }

    private updateWithAzureProperty = () => {
        this.props.UpdateSPUserWithAzureProps(this.state.data);
    }

    public render(): JSX.Element {
        const { data, propertyMappings } = this.state;
        return (
            <div>
                {(data && data.length > 0) ? (
                    <>
                        <EditableTable onRowDel={this.handleRowDel.bind(this)} data={data} propertyMappings={propertyMappings} isReadOnly={true} />
                        <div style={{ marginTop: "5px" }}>
                            <PrimaryButton text={strings.BtnUpdateUserProps} onClick={this.updateWithAzureProperty} style={{ marginRight: '5px' }} disabled={this.props.showProgress} />
                            {this.props.showProgress && <Spinner className={styles.generateTemplateLoader} label={strings.PropsUpdateLoader} ariaLive="assertive" labelPosition="right" />}
                        </div>
                    </>
                ) : (
                        <div><MessageContainer MessageScope={MessageScope.Info} Message={strings.EmptyTable} /></div>
                    )
                }
            </div>
        );
    }
}