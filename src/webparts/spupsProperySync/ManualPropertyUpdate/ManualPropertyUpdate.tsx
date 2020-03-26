import * as React from 'react';

import SearchBar from './SearchBar';
import EditableTable from './EditableTable';

export interface IManualPropertyUpdateProps {
    userProperties: any;
}

export interface IManualPropertyUpdateState {
    filterText: string;
    data: any;
}

export default class ManualPropertyUpdate extends React.Component<IManualPropertyUpdateProps, IManualPropertyUpdateState> {
    constructor(props: IManualPropertyUpdateProps) {
        super(props);
        this.state = {
            filterText: "",
            data: []
        };
    }

    public componentDidMount = async () => {
        this.setState({ data: this.props.userProperties });
    }

    public componentDidUpdate = (prevProps: IManualPropertyUpdateProps) => {
        if (prevProps.userProperties !== this.props.userProperties) {
            this.setState({ data: this.props.userProperties });
        }
    }

    public handleRowDel = (item) => {
        var index = this.state.data.indexOf(item);
        this.state.data.splice(index, 1);
        this.setState(this.state.data);
    }

    public handleProductTable = (evt) => {
        var newProp = {
            id: evt.target.id,
            name: evt.target.name,
            value: evt.target.value
        };
        var upProperties = this.state.data.slice();
        var newitem = upProperties.map((item) => {
            for (var key in item) {
                if (key == newProp.name && item.UserID == newProp.id) {
                    item[key] = newProp.value;
                }
            }
            return item;
        });
        this.setState({ data: newitem });
    }

    public render(): JSX.Element {
        const { filterText, data } = this.state;
        return (
            <div>
                {data && data.length > 0 &&
                    <>
                        <EditableTable onTableUpdate={this.handleProductTable.bind(this)} onRowDel={this.handleRowDel.bind(this)}
                            data={data} filterText={filterText} />
                    </>
                }
            </div>
        );
    }
}