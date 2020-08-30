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
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { IPersonaSharedProps, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { useDropzone } from 'react-dropzone';
import ImageResize from 'image-resize';

const imgResize_48 = new ImageResize({ format: 'png', width: 48, height: 48, output: 'base64' });
const imgResize_96 = new ImageResize({ format: 'png', width: 96, height: 96, output: 'base64' });
const imgResize_240 = new ImageResize({ format: 'png', width: 240, height: 240, output: 'base64' });

export interface IBulkPhotoSyncProps {

}

const BulkPhotoSync: React.FC<IBulkPhotoSyncProps> = (props) => {
    const appContext: AppContextProps = useContext(AppContext);
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
    /**
     * Build columns for Datalist.     
     */
    const _buildColumns = () => {
        let cols: IColumn[] = [];
        let col: string = 'path';
        cols.push({
            key: 'loginname', name: 'User ID', fieldName: col, minWidth: 250, maxWidth: 350,
            onRender: (item: any) => {
                return (<span>{item[col].replace('.' + item[col].split('.').pop(), '')}</span>);
            }
        } as IColumn);
        cols.push({
            key: 'preview', name: 'Photo', fieldName: col, minWidth: 100, maxWidth: 100,
            onRender: (item: any, index: number, column: IColumn) => {
                const authorPersona: IPersonaSharedProps = {
                    imageUrl: URL.createObjectURL(item),
                };
                return (
                    <div><Persona {...authorPersona} size={PersonaSize.large} /></div>
                );
            }
        } as IColumn);
        setColumns(cols);
    };
    const _listUploadedFiles = async () => {
        if (acceptedFiles.length > 0) {
            _buildColumns();
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
            {acceptedFiles && acceptedFiles.length > 0 &&
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