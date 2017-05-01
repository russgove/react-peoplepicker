import {
  IPersonaProps
} from 'office-ui-fabric-react';
import { Task, modes } from "../dataModel";
export interface IReactPeoplepickerExampleProps {
  description: string;
    peoplesearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>;
      ensureUser: (email) => Promise<any>;
      task:Task;
      mode:modes;

}
