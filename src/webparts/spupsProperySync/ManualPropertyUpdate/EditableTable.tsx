import * as React from 'react';
import DataRow from './DataRow';

export interface IEditableTableProps {
    onTableUpdate: () => void;
    onRowDel: () => void;
    onRowAdd: () => void;
    filterText: string;
    data: any[];
}

export default function EditableTable(props: IEditableTableProps) {
    console.log(props.data);
    var rowitem: any = props.data.map(function (item) {
        if (item.UserID.indexOf(props.filterText) === -1) {
            return;
        }        
        return (<DataRow item={item} onTableUpdate={props.onTableUpdate} onDelRow={props.onRowDel} key={item.UserID} />)
    });
    var keyData = JSON.parse(JSON.stringify(props.data));
    console.log(keyData);
    return (
        <div>
            <button type="button" onClick={props.onRowAdd} className="btn btn-success pull-right">Add</button>
            <table className="table table-bordered">
                <thead>
                    <tr>
                        {Object.keys(keyData[0]).map(key => {
                            return (<th>{key}</th>)
                        })}
                        {/* <th>Name</th>
                        <th>price</th>
                        <th>quantity</th>
                        <th>category</th> */}
                    </tr>
                </thead>
                <tbody>
                    {rowitem}
                </tbody>
            </table>
        </div>
    );
}