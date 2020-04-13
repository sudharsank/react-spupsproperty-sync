import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import * as strings from 'SpupsProperySyncWebPartStrings';
import { DetailsList, buildColumns, IColumn, DetailsListLayoutMode, ConstrainMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { ActionButton, IIconProps } from 'office-ui-fabric-react';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { css } from 'office-ui-fabric-react/lib';
import SPHelper from '../../../Common/SPHelper';
import * as moment from 'moment';
import MessageContainer from './MessageContainer';
import { MessageScope } from '../../../Common/IModel';
import SyncJobResults from './SyncJobResults';
import { orderBy } from 'lodash';

export interface ISyncJobsProps {
    helper: SPHelper;
}

export default function SyncJobsView(props: ISyncJobsProps) {
    const actionIcon: IIconProps = { iconName: 'InfoSolid' };
    const [loading, setLoading] = React.useState<boolean>(true);
    const [jobs, setJobs] = React.useState<any[]>([]);
    const [columns, setColumns] = React.useState<IColumn[]>([]);
    const [jobresults, setJobResults] = React.useState<string>('');
    const [hideDialog, setHideDialog] = React.useState<boolean>(true);

    const actionClick = (data) => {
        setJobResults(data.SyncResults);
        setHideDialog(false);
    };
    const StatusRender = (childprops) => {
        switch (childprops.Status.toLowerCase()) {
            case 'submitted':
                return (<div className={css(styles.status, styles.blue)}><Icon iconName="Save" /> {childprops.Status}</div>);
            case 'in-progress':
                return (<div className={css(styles.status, styles.orange)}><Icon iconName="ProgressRingDots" /> {childprops.Status}</div>);
            case 'completed':
                return (<div className={css(styles.status, styles.green)}><Icon iconName="Completed" /> {childprops.Status}</div>);
            case 'error':
            case 'completed with error':
                return (<div className={css(styles.status, styles.red)}><Icon iconName="ErrorBadge" /> {childprops.Status}</div>);
        }
    };
    const ActionRender = (actionProps) => {
        return (
            <ActionButton iconProps={actionIcon} allowDisabledFocus onClick={() => { actionClick(actionProps); }} />
        );
    };
    const _buildColumns = () => {
        let cols: IColumn[] = [];
        cols.push({ key: 'ID', name: 'ID', fieldName: 'ID', minWidth: 50, maxWidth: 50 } as IColumn);
        cols.push({ key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 250, maxWidth: 250 } as IColumn);
        cols.push({ key: 'SyncType', name: 'Sync Type', fieldName: 'SyncType', minWidth: 150, maxWidth: 150 } as IColumn);
        cols.push({
            key: 'Author', name: 'Author', fieldName: 'Author.Title', minWidth: 250, maxWidth: 250,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<div>{item.Author["Title"]}</div>);
            }
        } as IColumn);
        cols.push({
            key: 'Created', name: 'Created', fieldName: 'Created', minWidth: 150, maxWidth: 150,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<div>{moment(item.Created).format("DD, MMM YYYY hh:mm A")}</div>);
            }
        } as IColumn);
        cols.push({
            key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 200, maxWidth: 200,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<StatusRender Status={item.Status} />);
            }
        } as IColumn);
        cols.push({
            key: 'Actions', name: 'Actions', fieldName: 'ID', minWidth: 100, maxWidth: 100,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<ActionRender SyncResults={item.SyncedData} />);
            }
        });
        setColumns(cols);
    };
    const _buildJobsList = async () => {
        _buildColumns();
        let jobslist = await props.helper.getAllJobs();
        jobslist = orderBy(jobslist, ['ID'], ['desc']);
        setJobs(jobslist);
        setLoading(false);
    };
    const _closeDialog = () => {
        setHideDialog(true);
    };

    React.useEffect(() => {
        _buildJobsList();
    }, []);


    return (
        <div className={styles.syncjobsContainer}>
            {loading &&
                <ProgressIndicator label={strings.PropsLoader} description={strings.JobsListLoaderDesc} />
            }
            {(!loading && jobs && jobs.length > 0) ? (
                <DetailsList
                    items={jobs}
                    setKey="set"
                    columns={columns}
                    compact={true}
                    layoutMode={DetailsListLayoutMode.justified}
                    constrainMode={ConstrainMode.unconstrained}
                    isHeaderVisible={true}
                    selectionMode={SelectionMode.none}
                    enableShimmer={true}
                    className={styles.uppropertylist} />
            ) : (
                    <>
                        {!loading &&
                            <MessageContainer MessageScope={MessageScope.Info} Message={strings.EmptyTable} />
                        }
                    </>
                )}
            <Dialog hidden={hideDialog} onDismiss={_closeDialog} maxWidth='700'
                dialogContentProps={{
                    type: DialogType.close,
                    title: `${strings.JobResultsDialogTitle}`
                }}
                modalProps={{
                    isBlocking: true,
                    isDarkOverlay: true,
                    styles: { main: { maxHeight: 700 } },
                }}>
                <SyncJobResults helper={props.helper} data={jobresults} />
            </Dialog>

        </div>
    );
}

