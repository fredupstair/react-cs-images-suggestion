import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
 
import ReactCSsuggestion from './components/ReactCSsuggestion';
import { IReactCSsuggestionProps } from './components/IReactCSsuggestionProps';

export interface IReactCSsuggestionWebPartProps {
  description: string;
}

export default class ReactCSsuggestionWebPart extends BaseClientSideWebPart<IReactCSsuggestionWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactCSsuggestionProps> = React.createElement(
      ReactCSsuggestion,
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
