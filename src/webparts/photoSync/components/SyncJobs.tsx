import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import { useBoolean } from '@uifabric/react-hooks';
import styles from './PhotoSync.module.scss';
import * as strings from 'PhotoSyncWebPartStrings';
import { AppContext, AppContextProps } from '../common/AppContext';
import MessageContainer from '../common/MessageContainer';
import { MessageScope } from '../common/IModel';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { IPersonaSharedProps, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import * as moment from 'moment';
import { css } from 'office-ui-fabric-react/lib/Utilities';

const orderBy: any = require('lodash/orderBy');
const filter: any = require('lodash/filter');

export interface ISyncJobsProps {
    dateFormat: string;
}

const SyncJobs: React.FC<ISyncJobsProps> = (props) => {
    const appContext: AppContextProps = useContext(AppContext);
    const [loading, { setTrue: showLoading, setFalse: hideLoading }] = useBoolean(true);
    const [jobs, setJobs] = useState<any[]>([]);
    const [columns, setColumns] = React.useState<IColumn[]>([]);
    const [filItems, setFilItems] = useState<any[]>([]);
    const [searchText, setSearchText] = useState<string>('');

    const SyncTypeRender = (childprops) => {
        switch (childprops.SyncType.toLowerCase()) {
            case 'manual':
                return (
                    <div className={css(styles.fieldContent, styles.purplebgColor)}>
                        <span className={css(styles.spnContent, styles.purpleBox)}>{childprops.SyncType}</span>
                    </div>
                );
            case 'bulk':
                return (
                    <div className={css(styles.fieldContent, styles.yellowbgColor)}>
                        <span className={css(styles.spnContent, styles.yellowBox)}>{childprops.SyncType}</span>
                    </div>
                );
        }
    };

    const StatusRender = (childprops) => {
        switch (childprops.Status.toLowerCase()) {
            case 'submitted':
                return (
                    <div className={css(styles.fieldContent, styles.bluebgColor)}>
                        <span className={css(styles.spnContent, styles.blueBox)}>{childprops.Status}</span>
                    </div>
                );
            case 'in-progress':
                return (
                    <div className={css(styles.fieldContent, styles.orangebgColor)}>
                        <span className={css(styles.spnContent, styles.orangeBox)}>{childprops.Status}</span>
                    </div>
                );
            case 'completed':
                return (
                    <div className={css(styles.fieldContent, styles.greenbgColor)}>
                        <span className={css(styles.spnContent, styles.greenBox)}>{childprops.Status}</span>
                    </div>
                );
            case 'error':
            case 'completed with error':
                return (
                    <div className={css(styles.fieldContent, styles.redbgColor)}>
                        <span className={css(styles.spnContent, styles.redBox)}>{childprops.Status}</span>
                    </div>
                );
        }
    };

    const _buildColumns = () => {
        let cols: IColumn[] = [];
        cols.push({
            key: 'ID', name: 'ID', fieldName: 'ID', minWidth: 50, maxWidth: 50,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<div className={styles.fieldCustomizer}>{item.ID}</div>);
            }
        } as IColumn);
        cols.push({
            key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 100, maxWidth: 150,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<div className={styles.fieldCustomizer}>{item.Title}</div>);
            }
        } as IColumn);
        cols.push({
            key: 'SyncType', name: 'Sync Type', fieldName: 'SyncType', minWidth: 100, maxWidth: 100,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<div className={styles.fieldCustomizer}><SyncTypeRender SyncType={item.SyncType} /></div>);
            }
        } as IColumn);
        cols.push({
            key: 'Author', name: 'Author', fieldName: 'Author.Title', minWidth: 250, maxWidth: 250,
            onRender: (item: any, index: number, column: IColumn) => {
                const authorPersona: IPersonaSharedProps = {
                    imageUrl: `/_layouts/15/userphoto.aspx?Size=S&Username=${item["Author"].EMail}`,
                    text: item["Author"].Title,
                    className: styles.divPersona
                };
                return (
                    <div className={styles.fieldCustomizer}><Persona {...authorPersona} size={PersonaSize.size24} /></div>
                );
            }
        } as IColumn);
        cols.push({
            key: 'Created', name: 'Created', fieldName: 'Created', minWidth: 100, maxWidth: 100,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<div className={styles.fieldCustomizer}>{moment(item.Created).format(props.dateFormat ? props.dateFormat : 'DD/MM/YYYY')}</div>);
            }
        } as IColumn);
        cols.push({
            key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 150,
            onRender: (item: any, index: number, column: IColumn) => {
                return (<div className={styles.fieldCustomizer}><StatusRender Status={item.Status} /></div>);
            }
        } as IColumn);
        cols.push({
            key: 'Actions', name: 'Actions', fieldName: 'ID', minWidth: 100, maxWidth: 100,
            onRender: (item: any, index: number, column: IColumn) => {
                let disabled: boolean = ((item.Status.toLowerCase() == "error" && item.ErrorMessage && item.ErrorMessage.length > 0) || item.Status.toLowerCase().indexOf('completed') >= 0) ? false : true;
                //return (<ActionRender SyncResults={item.SyncedData} ErrorMessage={item.ErrorMessage} disabled={disabled} />);
            }
        });
        setColumns(cols);
    };

    const _onChangeSearchBox = (srchkey: string) => {
        setSearchText(srchkey);
        if (srchkey && srchkey.length > 0) {
            let filtered: any[] = filter(jobs, (o) => {
                return o.ID.toString().indexOf(srchkey.toLowerCase()) > -1 ||
                    o.Title.toLowerCase().indexOf(srchkey.toLowerCase()) > -1 || o['Author'].Title.toLowerCase().indexOf(srchkey.toLowerCase()) > -1 ||
                    o.Status.toLowerCase().indexOf(srchkey.toLowerCase()) > -1 || o.SyncType.toLowerCase().indexOf(srchkey.toLowerCase()) > -1;
            });
            setFilItems(filtered);
        } else setFilItems(jobs);
    };

    const _loadJobsList = async () => {
        let jobsList: any[] = await appContext.helper.getAllJobs();
        jobsList = orderBy(jobsList, ['ID'], ['desc']);
        console.log(jobsList);
        setJobs(jobsList);
        setFilItems(jobsList);
    };

    const _buildJobsList = async () => {
        _buildColumns();
        await _loadJobsList();
        hideLoading();
    };

    useEffect(() => {
        _buildJobsList();
    }, []);

    return (
        <div>
            {loading ? (
                <ProgressIndicator label={strings.PropsLoader} description={strings.JobsListLoaderDesc} />
            ) : (
                    <div className="ms-Grid-row" style={{ marginBottom: '5px', paddingLeft: '18px' }}>
                        <div>
                            <SearchBox
                                placeholder={`Search...`}
                                onChange={_onChangeSearchBox}
                                underlined={true}
                                iconProps={{ iconName: 'Filter' }}
                                value={searchText}
                                autoFocus={false}
                            //className={styles.favSearch}
                            />
                        </div>
                        {filItems && filItems.length > 0 ? (
                            <div style={{ overflowX: 'auto' }}>
                                <DetailsList
                                    items={filItems}
                                    setKey="set"
                                    columns={columns}
                                    compact={true}
                                    layoutMode={DetailsListLayoutMode.justified}
                                    constrainMode={ConstrainMode.unconstrained}
                                    isHeaderVisible={true}
                                    selectionMode={SelectionMode.none}
                                    enableShimmer={true}
                                //className={styles.detailsList}
                                />
                            </div>

                        ) : (
                                <MessageContainer MessageScope={MessageScope.Info} Message={strings.EmptyTable} />
                            )}
                    </div>
                )}
        </div>
    );
};

export default SyncJobs;