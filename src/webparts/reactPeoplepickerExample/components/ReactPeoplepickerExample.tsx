import * as React from 'react';
import {
  NormalPeoplePicker, CompactPeoplePicker, IBasePickerSuggestionsProps,
} from 'office-ui-fabric-react/lib/Pickers';
import { IReactPeoplepickerExampleProps } from './IReactPeoplepickerExampleProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
export default class ReactPeoplepickerExample extends React.Component<IReactPeoplepickerExampleProps, void> {
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
  public requestorChanged(req: Array<IPersonaProps>) { // need to call ensure user
    this.state.isDirty = true;
    if (req.length > 0) {
      console.log("requestor changedd " + req[0].optionalText);// I am only adding a single user. req[0] , others are ignored
      const email = req[0].optionalText;
      this.props.ensureUser(email).then((user) => {
        this.state.tr.RequestorId = user.data.Id;
        this.setState(this.state);
      }).catch((error) => {
        this.state.errorMessages.push(new md.Message(error.data.responseBody['odata.error'].message.value));
        this.setState(this.state);
      });
    }
    else {
      console.log("requestor removed ");// I am only adding a single user. req[0] , others are ignored
      this.state.tr.RequestorId = null;
    }
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
                defaultSelectedItems={this.state.tr.RequestorId ? [{ id: this.state.tr.RequestorId.toString(), primaryText: this.state.tr.RequestorName }] : []}
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
