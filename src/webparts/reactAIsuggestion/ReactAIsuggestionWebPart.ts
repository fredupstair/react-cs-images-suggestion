import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
 
import ReactAIsuggestion from './components/ReactAIsuggestion';
import { IReactAIsuggestionProps } from './components/IReactAIsuggestionProps';

export interface IReactAIsuggestionWebPartProps {
  description: string;
}

export default class ReactAIsuggestionWebPart extends BaseClientSideWebPart<IReactAIsuggestionWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactAIsuggestionProps> = React.createElement(
      ReactAIsuggestion,
      {
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

}
