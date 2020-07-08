import * as React from 'react';
import {useEffect, useState, useContext} from 'react';
import * as strings from 'PhotoSyncWebPartStrings';
import { AppContext, AppContextProps } from '../common/AppContext';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export interface IUserSelectionSyncProps {

}

const UserSelectionSync: React.FunctionComponent<IUserSelectionSyncProps> = (props) => {
    const appContext: AppContextProps = useContext(AppContext);
    const _selectedItems = (items: any[]) => {
        console.log(items);
    }
    return (
        <div>
            <PeoplePicker
                //disabled={disablePropsButtons || updatePropsLoader_Manual || updatePropsLoader_Azure}
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
        </div>
    );
};

export default UserSelectionSync;