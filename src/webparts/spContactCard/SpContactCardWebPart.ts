import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Log } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpContactCardWebPartStrings';
import SPFxPeopleCard, { IPeopleCardProps } from './components/SPFXPeopleCard/SPFxPeopleCard';
import { PersonaSize, PersonaInitialsColor } from 'office-ui-fabric-react';

export interface ISpContactCardWebPartProps {
  description: string;
}

export default class SpContactCardWebPart extends BaseClientSideWebPart<ISpContactCardWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPeopleCardProps> = React.createElement(
      SPFxPeopleCard, {  
        primaryText: this.context.pageContext.user.displayName,
        email: this.context.pageContext.user.email,
        serviceScope: this.context.serviceScope,
        class: 'persona-card',
        size: PersonaSize.extraLarge,
        initialsColor: PersonaInitialsColor.darkBlue,
        onCardOpenCallback: ()=>{
          console.log('WebPart','on card open callaback');
        },
        onCardCloseCallback: ()=>{
          console.log('WebPart','on card close callaback');
        }
      }
    );

    ReactDom.render(element, this.domElement);
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
}
