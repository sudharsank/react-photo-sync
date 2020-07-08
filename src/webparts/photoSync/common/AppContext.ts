import * as React from 'react';
import { IHelper } from './helper';
import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface AppContextProps {
    context: WebPartContext;
    helper: IHelper;
    displayMode: DisplayMode;
    openPropertyPane: () => void;
}

export const AppContext = React.createContext<AppContextProps>(undefined);