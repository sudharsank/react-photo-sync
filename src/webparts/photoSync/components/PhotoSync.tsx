import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './PhotoSync.module.scss';
import * as strings from 'PhotoSyncWebPartStrings';
import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { escape } from '@microsoft/sp-lodash-subset';
import { AppContext, AppContextProps } from '../common/AppContext';
import Helper, { IHelper } from '../common/helper';
import ConfigPlaceholder from '../common/ConfigPlaceholder';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/propertyFields/peoplePicker';


export interface IPhotoSyncProps {
    helper: IHelper;
    displayMode: DisplayMode;
    useFullWidth: boolean;
    appTitle: string;
    updateProperty: (value: string) => void;
    //templateLib: string;
    // AzFuncUrl: string;
    // UseCert: boolean;
    // dateFormat: string;
    allowedUsers: IPropertyFieldGroupOrPerson[];
    openPropertyPane: () => void;
}

const PhotoSync: React.FunctionComponent<IPhotoSyncProps> = (props) => {
    const [loading, setLoading] = useState<boolean>(true);
    const parentCtxValues: AppContextProps = {
        helper: props.helper,
        displayMode: props.displayMode,
        openPropertyPane: props.openPropertyPane
    }
    const showConfig = false; //!props.templateLib || !props.AzFuncUrl ? true : false;
    const _useFullWidth = () => {
        const jQuery: any = require('jquery');
        if (props.useFullWidth) {
            jQuery("#workbenchPageContent").prop("style", "max-width: none");
            jQuery(".SPCanvas-canvas").prop("style", "max-width: none");
            jQuery(".CanvasZone").prop("style", "max-width: none");
        } else {
            jQuery("#workbenchPageContent").prop("style", "max-width: 924px");
        }
    };

    useEffect(() => {
        _useFullWidth();
    }, [props.useFullWidth]);

    useEffect(() => {

    }, [])

    return (
        <AppContext.Provider value={parentCtxValues}>
            <div className={styles.photoSync}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <WebPartTitle displayMode={props.displayMode} title={props.appTitle ? props.appTitle : strings.DefaultAppTitle} updateProperty={props.updateProperty} />
                            {showConfig ? (
                                <ConfigPlaceholder />
                            ) : (
                                    <>
                                        {loading ? (
                                            <ProgressIndicator label={strings.AccessCheckDesc} description={strings.PropsLoader} />
                                        ) : (
                                            <>
                                            </>
                                        )}
                                    </>
                                )}
                        </div>
                    </div>
                </div>
            </div>
        </AppContext.Provider>
    );
};

export default PhotoSync;
