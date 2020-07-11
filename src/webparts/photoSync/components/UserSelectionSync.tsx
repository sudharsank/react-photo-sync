import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import { useBoolean } from '@uifabric/react-hooks';
import * as strings from 'PhotoSyncWebPartStrings';
import { AppContext, AppContextProps } from '../common/AppContext';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import MessageContainer from '../common/MessageContainer';
import { MessageScope, IUserPickerInfo } from '../common/IModel';
import { PrimaryButton } from 'office-ui-fabric-react/lib/components/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { IPersonaSharedProps, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import styles from './PhotoSync.module.scss';
import { divProperties } from 'office-ui-fabric-react/lib/Utilities';

const filter: any = require('lodash/filter');

export interface IUserSelectionSyncProps {

}

const UserSelectionSync: React.FunctionComponent<IUserSelectionSyncProps> = (props) => {
    const appContext: AppContextProps = useContext(AppContext);
    const [selectedUsers, setSelectedUsers] = useState<any[]>([]);
    const [reloadGetProperties, setReloadGetProperties] = useState<boolean>(false);
    const [clearData, { toggle: toggleClearData, setFalse: hideClearData }] = useBoolean(false);
    const [showPhotoLoader, { toggle: togglePhotoLoader, setFalse: hidePhotoLoader }] = useBoolean(false);
    const [disableButton, { toggle: toggleDisableButton, setFalse: enableButton }] = useBoolean(false);
    const [disableUserPicker, { toggle: toggleDisableUserPicker }] = useBoolean(false);
    const [columns, setColumns] = useState<IColumn[]>([]);

    const _buildColumns = (colValues: string[]) => {
        let cols: IColumn[] = [];
        colValues.map(col => {
            if (col.toLowerCase() == "title") {
                cols.push({
                    key: 'title', name: 'Title', fieldName: col, minWidth: 150, maxWidth: 200,
                } as IColumn);
            }
            if (col.toLowerCase() == "loginname") {
                cols.push({
                    key: 'loginname', name: 'User ID', fieldName: col, minWidth: 250, maxWidth: 350,
                    onRender: (item: any) => {
                        return (<span>{item[col].split('|')[2]}</span>)
                    }
                } as IColumn);
            }
            if (col.toLowerCase() == "photourl") {
                cols.push({
                    key: 'photourl', name: 'SP Profile Photo', fieldName: col, minWidth: 100, maxWidth: 100,
                    onRender: (item: any, index: number, column: IColumn) => {
                        const authorPersona: IPersonaSharedProps = {
                            imageUrl: item[col],
                        };
                        return (
                            <div><Persona {...authorPersona} size={PersonaSize.large} /></div>
                        );
                    }
                } as IColumn);
            }
            if (col.toLowerCase() == "aadphotourl") {
                cols.push({
                    key: 'aadphotourl', name: 'Azure Profile Photo', fieldName: col, minWidth: 100, maxWidth: 100,
                    onRender: (item: any, index: number, column: IColumn) => {
                        if (item[col]) {
                            const authorPersona: IPersonaSharedProps = {
                                imageUrl: item[col],
                            };
                            return (
                                <div><Persona {...authorPersona} size={PersonaSize.large} /></div>
                            );
                        } else return (<></>);
                    }
                } as IColumn);
            }
        });
        setColumns(cols);
    };
    const _selectedItems = (items: any[]) => {
        let userInfo: IUserPickerInfo[] = [];
        if (items && items.length > 0) {
            items.map(item => {
                userInfo.push({
                    Title: item.text,
                    LoginName: item.loginName,
                    PhotoUrl: item.imageUrl
                });
            });
        }
        setSelectedUsers(userInfo);
        _buildColumns(Object.keys(userInfo[0]));
        enableButton();
    };
    const _getPhotosFromAzure = async () => {
        toggleDisableUserPicker();
        toggleDisableButton();
        togglePhotoLoader();
        let res: any[] = await appContext.helper.getUserPhotoFromAADForDisplay(selectedUsers);
        if (res && res.length > 0) {
            let tempUsers: IUserPickerInfo[] = selectedUsers;
            res.map(response => {
                if (response.responses && response.responses.length > 0) {
                    response.responses.map(finres => {
                        var fil = filter(tempUsers, (o) => { return o.LoginName == finres.id });
                        if (fil && fil.length > 0) {
                            fil[0].AADPhotoUrl = finres.body.error ? '' : "data:image/jpg;base64," + finres.body;
                        }
                    });
                }
            });
            console.log(tempUsers);
            setSelectedUsers(tempUsers);
            _buildColumns(Object.keys(tempUsers[0]));
        }
        //toggleDisableButton();
        toggleDisableUserPicker();
        togglePhotoLoader();
    };
    return (
        <div>
            <PeoplePicker
                disabled={disableUserPicker}
                context={appContext.context}
                titleText={strings.PPLPickerTitleText}
                personSelectionLimit={10}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={false}
                isRequired={false}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={500}
                selectedItems={_selectedItems}
            //defaultSelectedUsers={selectedUsers.length > 0 ? this._getSelectedUsersLoginNames(selectedUsers) : []}
            />
            {/* {reloadGetProperties ? (
                <>
                    {selectedUsers.length > 0 &&
                        <div>
                            <MessageContainer MessageScope={MessageScope.Info} Message={strings.Photo_UserListChanges} />
                        </div>
                    }
                    {selectedUsers.length <= 0 && !clearData &&
                        <div>
                            <MessageContainer MessageScope={MessageScope.Info} Message={strings.Photo_UserListEmpty} ShowDismiss={true} />
                        </div>
                    }
                </>
            ) : (
                    <></>
                )
            } */}
            {selectedUsers && selectedUsers.length > 0 &&
                <>
                    <div style={{ marginTop: "5px" }}>
                        <PrimaryButton text={strings.BtnAzurePhotoProps} onClick={_getPhotosFromAzure} disabled={disableButton} />
                        {showPhotoLoader && <Spinner className={styles.generateTemplateLoader} label={strings.PropsLoader} ariaLive="assertive" labelPosition="right" />}
                    </div>
                    <div style={{ marginTop: '5px' }}>
                        <DetailsList
                            items={selectedUsers}
                            setKey="set"
                            columns={columns}
                            compact={true}
                            layoutMode={DetailsListLayoutMode.justified}
                            constrainMode={ConstrainMode.unconstrained}
                            isHeaderVisible={true}
                            selectionMode={SelectionMode.none}
                            enableShimmer={true} />
                    </div>
                </>
            }
        </div>
    );
};

export default UserSelectionSync;