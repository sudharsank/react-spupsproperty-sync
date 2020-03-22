import * as React from 'react';

export interface IEditableCellProps {
    cellData: any;
    onTableUpdate: () => void;
}

export default function EditableCell(props: IEditableCellProps) {
    return (
        <td>
            {!props.cellData.label ? (
                <input type='text' name={props.cellData.type} id={props.cellData.id} value={props.cellData.value} onChange={props.onTableUpdate} />
            ) : (
                <div>
                    <label>{props.cellData.value}</label>
                    <img src={props.cellData.ImageUrl} />
                </div>                
            )}            
        </td>
    );
}