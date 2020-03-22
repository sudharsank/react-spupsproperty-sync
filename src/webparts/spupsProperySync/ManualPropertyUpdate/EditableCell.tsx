import * as React from 'react';
import { render } from 'react-dom';

export interface IEditableCellProps {
    cellData: any;
    onTableUpdate: () => void;
}

export default function EditableCell(props: IEditableCellProps) {
    return (
        <td>
            <input type='text' name={props.cellData.type} id={props.cellData.id} value={props.cellData.value} onChange={props.onTableUpdate} />
        </td>
    );
}