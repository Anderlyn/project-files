import * as React from 'react';
import styles from './ProjectFilesViewer.module.scss';
import { IProjectFilesViewerProps } from './IProjectFilesViewerProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ProjectFilesViewer extends React.Component<IProjectFilesViewerProps, {}> {
  constructor(props){
    super(props);
    this.state = {

    };
  }
  public render(): React.ReactElement<{IProjectFilesViewerProps}> {
    return (
      <div className={ styles.projectFilesViewer }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
