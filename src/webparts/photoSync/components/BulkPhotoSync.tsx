import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import { useBoolean } from '@uifabric/react-hooks';
import styles from './PhotoSync.module.scss';
import * as strings from 'PhotoSyncWebPartStrings';
import { AppContext, AppContextProps } from '../common/AppContext';
import MessageContainer from '../common/MessageContainer';
import { MessageScope, IUserPickerInfo } from '../common/IModel';
import { PrimaryButton } from 'office-ui-fabric-react/lib/components/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { IPersonaSharedProps, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { useDropzone } from 'react-dropzone';
import ImageResize from 'image-resize';
import { css } from 'office-ui-fabric-react/lib/Utilities';

const imgResize_48 = new ImageResize({ format: 'png', width: 48, height: 48, output: 'base64' });
const imgResize_96 = new ImageResize({ format: 'png', width: 96, height: 96, output: 'base64' });
const imgResize_240 = new ImageResize({ format: 'png', width: 240, height: 240, output: 'base64' });

const map: any = require('lodash/map');
const find: any = require('lodash/find');

export interface IBulkPhotoSyncProps {

}

const BulkPhotoSync: React.FC<IBulkPhotoSyncProps> = (props) => {
    const appContext: AppContextProps = useContext(AppContext);
    const [loading, { setTrue: showLoading, setFalse: hideLoading }] = useBoolean(false);
    const [columns, setColumns] = useState<IColumn[]>([]);
    const { getRootProps, getInputProps, fileRejections, acceptedFiles } = useDropzone({
        accept: 'image/jpeg, image/jpg, image/png',
        // onDrop: acceptedFiles => {
        //     setFiles(acceptedFiles.map(file => {
        //         file['preview'] = URL.createObjectURL(file);
        //         return file;
        //     }));
        //     console.log(files);
        // }
    });
    // const thumbs = files.map(file => (
    //     <div className={styles.thumb} key={file.name}>
    //         <div className={styles.thumbInner}>
    //             <img src={file.preview} />
    //         </div>
    //     </div>
    // ));

    const StatusRender = (childprops) => {
        switch (childprops.Status.toLowerCase()) {
            case 'success':
                return (
                    <div className={css(styles.fieldContent, styles.greenbgColor)}>
                        <span className={css(styles.spnContent, styles.greenBox)}>{childprops.Status}</span>
                    </div>
                );
            case 'failed':
                return (
                    <div className={css(styles.fieldContent, styles.redbgColor)}>
                        <span className={css(styles.spnContent, styles.redBox)}>{childprops.Status}</span>
                    </div>
                );
        }
    };

    /**
     * Build columns for Datalist.     
     */
    const _buildColumns = () => {
        let cols: IColumn[] = [];
        let col: string = 'path';
        cols.push({
            key: 'loginname', name: 'User ID', fieldName: col, minWidth: 250, maxWidth: 350,
            onRender: (item: any) => {
                return (<div className={styles.fieldCustomizer}>{item[col].replace('.' + item[col].split('.').pop(), '')}</div>);
            }
        } as IColumn);
        cols.push({
            key: 'usertitle', name: 'Title', fieldName: 'title', minWidth: 250, maxWidth: 350,
            onRender: (item: any, index: number, column: IColumn) => {
                const authorPersona: IPersonaSharedProps = {
                    imageUrl: `/_layouts/15/userphoto.aspx?Size=S&username=${item.name.replace('.' + item.name.split('.').pop(), '')}`,
                    text: item.title,
                    className: styles.divPersona
                };
                return (
                    <div className={styles.fieldCustomizer}><Persona {...authorPersona} size={PersonaSize.size24} /></div>
                );
            }
        } as IColumn);
        cols.push({
            key: 'preview', name: 'Photo', fieldName: col, minWidth: 100, maxWidth: 100,
            onRender: (item: any, index: number, column: IColumn) => {
                // const authorPersona: IPersonaSharedProps = {
                //     imageUrl: URL.createObjectURL(item),
                // };
                // return (
                //     <div><Persona {...authorPersona} size={PersonaSize.large} /></div>
                // );
                return (
                    <div className={styles.fieldCustomizer}>
                        <img style={{ width: '50px' }} src={URL.createObjectURL(item)} />
                    </div>
                )
            }
        } as IColumn);
        cols.push({
            key: 'status', name: 'Status', fieldName: 'status', minWidth: 250, maxWidth: 350,
            onRender: (item: any) => {
                return (<div className={styles.fieldCustomizer}><StatusRender Status={item.status} /></div>);
            }
        } as IColumn);
        setColumns(cols);
    };
    const _listUploadedFiles = async () => {
        if (acceptedFiles.length > 0) {
            showLoading();
            let userids: string[] = map(acceptedFiles, (o) => { return o.name.replace('.' + o.name.split('.').pop(), ''); });
            let userinfo: any[] = await appContext.helper.getUsersInfo(userids);
            if (userinfo && userinfo.length > 0) {
                userinfo.map((user: any) => {
                    let fil: any = find(acceptedFiles, (o) => { return o.name.replace('.' + o.name.split('.').pop(), '') == user.loginname; });
                    if (fil) {
                        fil['title'] = user.title;
                        fil['status'] = user.status;
                    }
                });
            }
            _buildColumns();
            hideLoading();
        }
    };
    useEffect(() => {
        _listUploadedFiles();
    }, [acceptedFiles]);
    // useEffect(() => () => {
    //     console.log('yes');
    //     // Make sure to revoke the data uris to avoid memory leaks
    //     files.forEach(file => URL.revokeObjectURL(file.preview));
    // }, [files]);
    return (
        <div>
            <div style={{ margin: '5px 0px' }}>
                <MessageContainer MessageScope={MessageScope.Info} Message={strings.BulkSyncNote} />
            </div>
            <section className={styles.dropZoneContainer}>
                <div {...getRootProps({ className: styles.dropzone })}>
                    <input {...getInputProps()} />
                    <p>{strings.BulkPhotoDragDrop}</p>
                </div>
                {/* <aside className={styles.thumbsContainer}>
                    {thumbs}
                </aside> */}
            </section>
            {loading &&
                <ProgressIndicator label="Loading Photos..." description="Please wait..." />
            }
            {!loading && acceptedFiles && acceptedFiles.length > 0 &&
                <div className={styles.detailsListContainer}>
                    <DetailsList
                        items={acceptedFiles}
                        setKey="set"
                        columns={columns}
                        compact={true}
                        layoutMode={DetailsListLayoutMode.justified}
                        constrainMode={ConstrainMode.unconstrained}
                        isHeaderVisible={true}
                        selectionMode={SelectionMode.none}
                        enableShimmer={true} />
                </div>
            }
        </div>
    );
};

export default BulkPhotoSync;