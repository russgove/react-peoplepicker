import {
  IPersonaProps
} from 'office-ui-fabric-react';
import { Task, modes } from "../dataModel";
export interface IReactPeoplepickerExampleProps {

  peoplesearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>;
  ensureUser: (email) => Promise<any>;
  task: Task;
  mode: modes;
  save: (tr) => any; //make this return promise
  cancel: () => any;

}
