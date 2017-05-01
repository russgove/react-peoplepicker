import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version ,UrlQueryParameterCollection} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SearchQuery, SearchResults, SortDirection } from "sp-pnp-js";
import * as pnp from "sp-pnp-js";
import * as _ from 'lodash';
import { Task, modes } from "./dataModel";
import * as strings from 'reactPeoplepickerExampleStrings';
import ReactPeoplepickerExample from './components/ReactPeoplepickerExample';
import { IReactPeoplepickerExampleProps } from './components/IReactPeoplepickerExampleProps';
import { IReactPeoplepickerExampleWebPartProps } from './IReactPeoplepickerExampleWebPartProps';
import {
  IPersonaProps, PersonaPresence
} from 'office-ui-fabric-react';
export default class ReactPeoplepickerExampleWebPart extends BaseClientSideWebPart<IReactPeoplepickerExampleWebPartProps> {
  private task: Task;
  private reactElement: React.ReactElement<IReactPeoplepickerExampleProps>;
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });
      //  this.loadData();
    });
  }
  public render(): void {

    // hide the ribbon
    if (document.getElementById("s4-ribbonrow")) {
      document.getElementById("s4-ribbonrow").style.display = "none";
    }

    let formProps: IReactPeoplepickerExampleProps = {

      ensureUser: this.ensureUser,
      mode: this.properties.mode,

      peoplesearch: this.peoplesearch,
      task: new Task(),
      save: this.save.bind(this)
    };
    let batch = pnp.sp.createBatch();
    // get loolup field values here using the natch
    var queryParameters = new UrlQueryParameterCollection(window.location.href);

    if (this.properties.mode !== modes.NEW) {
      if (queryParameters.getValue("Id")) {
        const id: number = parseInt(queryParameters.getValue("Id"));
        let fields = "*,ParentTR/Title,Requestor/Title";
        // get the requested tr
        pnp.sp.web.lists.getByTitle("Tasks").items.getById(id).expand("AssignedTo").select(fields).inBatch(batch).get()

          .then((item) => {
            formProps.task = new Task();
            this.moveFieldsToTR(formProps.task, item);


          })
          .catch((error) => {
            console.log("ERROR, An error occured fetching the listitem");
            console.log(error.message);

          });
        // get the Child trs

      }
    }
    else {
      console.log("ERROR, Id not specified with New or Edit form");
    }

    batch.execute().then((value) => {

      this.reactElement = React.createElement(ReactPeoplepickerExample, formProps);
      ReactDom.render(this.reactElement, this.domElement);
    }
    );
   
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
  public peoplesearch(searchText: string, currentSelected: IPersonaProps[]): Promise<IPersonaProps[]> {


    let sq: SearchQuery = {
      Querytext: "Title:" + searchText + "* ContentClass=urn:content-class:SPSPeople",
      SourceId: "b09a7990-05ea-4af9-81ef-edfab16c4e31",  //http://www.dotnetmafia.com/blogs/dotnettipoftheday/archive/2013/01/04/list-of-common-sharepoint-2013-result-source-ids.aspx
      RowLimit: 50,
      SelectProperties: ["PreferredName", "Department", "JobTitle", "PictureURL",
        "OfficeNumber", "WorkEmail"]
      ///SortList: [{ Property: "PreferredName", Direction: SortDirection.Ascending }] arghhh-- not sortable
      // SelectProperties: ["*"]
    };
    return pnp.sp.search(sq).then((results: SearchResults) => {
      let resultsPersonas: Array<IPersonaProps> = [];
      for (let element of results.PrimarySearchResults) {
        const temp = element as any;
        let personapprop: IPersonaProps = {
          primaryText: temp.PreferredName,
          secondaryText: temp.JobTitle,
          tertiaryText: (temp.OfficeNumber) ? temp.Department + "(" + temp.OfficeNumber + ") " : temp.Department,
          imageUrl: temp.PictureURL,
          imageInitials: temp.contentclass,
          presence: PersonaPresence.none,
          optionalText: temp.WorkEmail // need this for ensureuser

        };
        resultsPersonas.push(personapprop);
      }
      return _.sortBy(resultsPersonas, "primaryText");
    });


  }
  protected ensureUser(email): Promise<any> {
    return pnp.sp.web.ensureUser(email);
  }

}
