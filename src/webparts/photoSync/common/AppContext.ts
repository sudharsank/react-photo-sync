import * as React from 'react';
import { IHelper } from './helper';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface AppContextProps {
    helper: IHelper;
    displayMode: DisplayMode;
    openPropertyPane: () => void;
}

export const AppContext = React.createContext<AppContextProps>(undefined);