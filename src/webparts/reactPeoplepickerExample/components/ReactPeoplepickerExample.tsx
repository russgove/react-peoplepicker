import * as React from 'react';
import {
  NormalPeoplePicker, CompactPeoplePicker, IBasePickerSuggestionsProps,
} from 'office-ui-fabric-react/lib/Pickers';
import { IReactPeoplepickerExampleProps } from './IReactPeoplepickerExampleProps';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { escape } from '@microsoft/sp-lodash-subset';
import { Task, modes } from "../dataModel";
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import * as md from "./MessageDisplay";
import MessageDisplay from "./MessageDisplay";
export interface IReactPeoplepickerExampleState{
  isDirty:boolean;
  task:Task;
    errorMessages: Array<md.Message>;
}
export default class ReactPeoplepickerExample extends React.Component<IReactPeoplepickerExampleProps, IReactPeoplepickerExampleState> {
    constructor(props: IReactPeoplepickerExampleProps) {
    super(props);
    this.state = {
      task: props.task,
   errorMessages: [],
      isDirty: false,
    };
    this.SaveButton = this.SaveButton.bind(this);
    this.ModeDisplay = this.ModeDisplay.bind(this);
    this.StatusDisplay = this.StatusDisplay.bind(this);
    this.save = this.save.bind(this);
    this.cancel = this.cancel.bind(this);


  }
  public resolveSuggestions(searchText: string, currentSelected: IPersonaProps[]): Promise<IPersonaProps> | IPersonaProps[] {

    return this.props.peoplesearch(searchText, currentSelected);
  }
  public getTextFromItem(persona: IPersonaProps): string {

    return persona.primaryText;
  }
  public renderPeople(person: IPersonaProps): JSX.Element {

    return (<Persona
      size={PersonaSize.large}
      primaryText={person.primaryText}
      secondaryText={person.secondaryText}

      tertiaryText={person.tertiaryText}
      imageUrl={person.imageUrl}
      imageShouldFadeIn={true}

    />);

  }
  
  public SaveButton(): JSX.Element {
    if (this.props.mode === modes.DISPLAY) {
      return <div />;
    } else return (
      <span style={{ margin: 20 }}>
        <a href="#" onClick={this.save} style={{ border: 5, backgroundColor: 'lightBlue', fontSize: 'large' }}>
          <i className="ms-Icon ms-Icon--Save"></i>Save
        </a>
      </span>
    );
  }
  public requestorChanged(req: Array<IPersonaProps>) { // need to call ensure user
    this.state.isDirty = true;
    if (req.length > 0) {
      console.log("requestor changedd " + req[0].optionalText);// I am only adding a single user. req[0] , others are ignored
      const email = req[0].optionalText;
      this.props.ensureUser(email).then((user) => {
        this.state.task.AssignedToId = user.data.Id;
        this.setState(this.state);
      }).catch((error) => {
        this.state.errorMessages.push(new md.Message(error.data.responseBody['odata.error'].message.value));
        this.setState(this.state);
      });
    }
    else {
      console.log("requestor removed ");// I am only adding a single user. req[0] , others are ignored
      this.state.task.AssignedToId = null;
    }
  }
    public ModeDisplay(): JSX.Element {
    return (
      <Label>MODE : {modes[this.props.mode]}</Label>
    );

  }
    public cancel() {

    this.props.cancel();
    return false; // stop postback

  }
  public StatusDisplay(): JSX.Element {
    return (
      <Label>Status : {(this.state.isDirty) ? "Unsaved" : "Saved"}</Label>
    );

  }
   public isValid(): boolean {

    this.state.errorMessages = [];
    let errorsFound = false;
    if (!this.state.task.Title) {
      this.state.errorMessages.push(new md.Message("Request #  is required"));
      errorsFound = true;
    }
  
    if (!this.state.task.DueDate) {
      this.state.errorMessages.push(new md.Message("Due Date  is required"));
      errorsFound = true;
    }
  
    if (!this.state.task.Priority) {
      this.state.errorMessages.push(new md.Message("Proiority  is required"));
      errorsFound = true;
    }
    if (!this.state.task.Status) {
      this.state.errorMessages.push(new md.Message("Status  is required"));
      errorsFound = true;
    }
    if (!this.state.task.AssignedToId) {
      this.state.errorMessages.push(new md.Message("Requestor is required"));
      errorsFound = true;
    }


    return !errorsFound;
  }
    public save() {

  

    if (this.isValid()) {
      this.props.save(this.state.task)
        .then((result) => {
          this.state.isDirty = false;
          this.setState(this.state);
        })
        .catch((response) => {
          this.state.errorMessages.push(new md.Message(response.data.responseBody['odata.error'].message.value));
          this.setState(this.state);
        });
    } else {
      this.setState(this.state); // show errors
    }
    return false; // stop postback
  }
  public render(): React.ReactElement<IReactPeoplepickerExampleProps> {
    const suggestionProps: IBasePickerSuggestionsProps = {
      suggestionsHeaderText: 'Suggested People',
      noResultsFoundText: 'No results found',
      loadingText: 'Loading',

    };
    return (
      <div>
        <NormalPeoplePicker
          defaultSelectedItems={this.state.task.AssignedToId ? [{
             id: this.state.task.AssignedToId.toString(), 
             primaryText: this.state.task.AssignedTo
              }] : []}
          onResolveSuggestions={this.resolveSuggestions.bind(this)}
          pickerSuggestionsProps={suggestionProps}
          getTextFromItem={this.getTextFromItem}
          onRenderSuggestionsItem={this.renderPeople}
          onChange={this.requestorChanged.bind(this)}
        />
      </div>
    );
  }
}
