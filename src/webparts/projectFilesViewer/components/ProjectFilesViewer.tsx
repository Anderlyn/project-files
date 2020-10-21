import * as React from 'react';
import { IProjectFilesViewerProps } from './models/IProjectFilesViewerProps';
import { State} from './models/IProjectFilesViewerState';
import { escape } from '@microsoft/sp-lodash-subset';
export default class ProjectFilesViewer extends React.Component<IProjectFilesViewerProps, State> {
  constructor(props){
    super(props);
    this.state = {
      url: ""
    };
  }
  componentDidMount(){
    var id:string = document.URL.split('?')[1].split('&')[0].split('=')[1];
    var site:string = document.URL.toLowerCase().split('project')[0];
    var req:any = new XMLHttpRequest();
    req.open('GET', site + '_api/ProjectData/Projects?$format=json&$filter=ProjectId%20eq%20guid%27' + id + '%27&$select=ProjectWorkspaceInternalUrl,Fase');
    req.send('');
    req.onreadystatechange = function(){
      if(req.readyState === 4){
        console.log(req.response);
        var parsed = JSON.parse(req.response);
        var projSite = parsed.value[0].ProjectWorkspaceInternalUrl;
        var FaseProj = parsed.value[0].Fase;
        console.log(FaseProj);
        this.setState({
          url:'"'+ projSite +'"/Shared Documents/"'+ FaseProj +'"' 
        })
      }
    }
  }
  public render(): React.ReactElement<{IProjectFilesViewerProps}> {
    return (
      <div>
        <iframe src={this.state.url}></iframe>
      </div>
    );
  }
}
