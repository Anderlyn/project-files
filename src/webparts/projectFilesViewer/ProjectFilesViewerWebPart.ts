import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'ProjectFilesViewerWebPartStrings';
import ProjectFilesViewer from './components/ProjectFilesViewer';
import { IProjectFilesViewerProps } from './components/models/IProjectFilesViewerProps';

export interface IProjectFilesViewerWebPartProps {
  description: string;
}

export default class ProjectFilesViewerWebPart extends BaseClientSideWebPart <IProjectFilesViewerWebPartProps> {
  public async onInit():Promise<void>{
    const _ = await super.onInit();
    let canvas:HTMLElement = document.querySelector(".SPCanvas-canvas");
    canvas.style.maxWidth = "none";
    let canvasZone:HTMLElement = document.querySelector(".CanvasZone");
    canvasZone.style.maxWidth = "none";
    let spNavBar:HTMLElement = document.querySelector(".spNav_f7fd2212");
    spNavBar.style.maxWidth = "none";
  }

  public render(): void {
    this.loadSPJSOMScripts();
    const element: React.ReactElement<IProjectFilesViewerProps> = React.createElement(
      ProjectFilesViewer,
      {
        description: this.properties.description
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
  private loadSPJSOMScripts() {
    SPComponentLoader.loadScript('/_layouts/15/init.js', {
      globalExportsName: '$_global_init'
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', {
        globalExportsName: 'Sys'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/ScriptResx.ashx?name=sp.res&culture=en-us', {
        globalExportsName: 'Sys'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/SP.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/sp.init.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/sp.ui.dialog.js', {
        globalExportsName: 'SP'
      });
    });
  }
}
