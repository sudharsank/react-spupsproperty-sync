import * as React from 'react';
import DataRow from './DataRow';

export interface IEditableTableProps {
    onTableUpdate: () => void;
    onRowDel: () => void;
    filterText: string;
    data: any[];
}

export default function EditableTable(props: IEditableTableProps) {
    var keyData = JSON.parse(JSON.stringify(props.data));
    var columns = Object.keys(keyData[0]);
    var rowitem: any = props.data.map((item) => {
        return (<DataRow item={item} columns={columns} onTableUpdate={props.onTableUpdate} onDelRow={props.onRowDel} key={item.UserID} />);
    });
    return (
        <div style={{overflowX: 'auto'}}>
            <table className="table table-bordered">
                <thead>
                    <tr>
                        {columns.map(key => {
                            if (key !== "ImageUrl") {
                                return (<th>{key}</th>);
                            }
                        })}
                    </tr>
                </thead>
                <tbody>
                    {rowitem}
                </tbody>
            </table>
        </div>
    );
}