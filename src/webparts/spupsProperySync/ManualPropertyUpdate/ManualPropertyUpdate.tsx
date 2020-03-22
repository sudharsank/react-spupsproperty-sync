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
        }
    }

    public componentDidMount = async () => {
        // this.setState({
        //     data: [
        //         {
        //             id: 1,
        //             category: 'Sporting Goods',
        //             price: '49.99',
        //             qty: 12,
        //             name: 'football'
        //         }, {
        //             id: 2,
        //             category: 'Sporting Goods',
        //             price: '9.99',
        //             qty: 15,
        //             name: 'baseball'
        //         }, {
        //             id: 3,
        //             category: 'Sporting Goods',
        //             price: '29.99',
        //             qty: 14,
        //             name: 'basketball'
        //         }, {
        //             id: 4,
        //             category: 'Electronics',
        //             price: '99.99',
        //             qty: 34,
        //             name: 'iPod Touch'
        //         }, {
        //             id: 5,
        //             category: 'Electronics',
        //             price: '399.99',
        //             qty: 12,
        //             name: 'iPhone 5'
        //         }, {
        //             id: 6,
        //             category: 'Electronics',
        //             price: '199.99',
        //             qty: 23,
        //             name: 'nexus 7'
        //         }
        //     ]
        // });
        this.setState({ data: this.props.userProperties });
        console.log('ManualPropertyUpdate', this.props.userProperties);
    }

    componentDidUpdate = (prevProps: IManualPropertyUpdateProps) => {
        if (prevProps.userProperties !== this.props.userProperties) {
            this.setState({ data: this.props.userProperties });
        }
    }

    handleUserInput(filterText) {
        this.setState({ filterText: filterText });
    };

    handleRowDel(item) {
        var index = this.state.data.indexOf(item);
        this.state.data.splice(index, 1);
        this.setState(this.state.data);
    };

    handleAddEvent(evt) {
        var id = (+ new Date() + Math.floor(Math.random() * 999999)).toString(36);
        var item = {
            id: id,
            name: "",
            price: "",
            category: "",
            qty: 0
        }
        this.state.data.push(item);
        this.setState(this.state.data);
    }

    handleProductTable(evt) {
        var item = {
            id: evt.target.id,
            name: evt.target.name,
            value: evt.target.value
        };
        var products = this.state.data.slice();
        var newitem = products.map(function (item) {
            for (var key in item) {
                if (key == item.name && item.id == item.id) {
                    item[key] = item.value;
                }
            }
            return item;
        });
        this.setState({ data: newitem });
        //  console.log(this.state.products);
    };

    public render(): JSX.Element {
        const { filterText, data } = this.state;
        return (
            <div>
                {data && data.length > 0 &&
                    <>
                        <SearchBar onUserInput={this.handleUserInput.bind(this)} />
                        <EditableTable onTableUpdate={this.handleProductTable.bind(this)} onRowAdd={this.handleAddEvent.bind(this)} onRowDel={this.handleRowDel.bind(this)}
                            data={data} filterText={filterText} />
                    </>
                }
            </div>
        );
    }
}