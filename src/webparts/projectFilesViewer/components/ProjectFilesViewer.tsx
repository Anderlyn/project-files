import * as React from 'react';
import { IProjectFilesViewerProps } from './models/IProjectFilesViewerProps';
import { IProjectFilesViewerState} from './models/IProjectFilesViewerState';
import { Language } from "./models/Language"
import { escape } from '@microsoft/sp-lodash-subset';
export default class ProjectFilesViewer extends React.Component<IProjectFilesViewerProps, IProjectFilesViewerState> {
  constructor(props){
    super(props);
    this.state = {
      url: "",
      loading: true,
      isPDP: false
    };
  }
  componentDidMount(){
    let iProps:Language;
    if(this.props.language){
      switch(this.props.language.toLowerCase()){
        case "eng":
          iProps = {
            endpointPlace: "Projects",
            projectId: "ProjectId",
            urlProperty: "ProjectWorkspaceInternalUrl",
            documentPlace: "Shared Documents"
          }
        break;
        case "esp":
          iProps = {
            endpointPlace: "Proyectos",
            projectId: "IdDelProyecto",
            urlProperty: "URLInternaDelEspacioDeTrabajoDelProyecto",
            documentPlace: "Documentos compartidos"
          }
        break;
        default:
          iProps = {
            endpointPlace: "Proyectos",
            projectId: "IdDelProyecto",
            urlProperty: "URLInternaDelEspacioDeTrabajoDelProyecto",
            documentPlace: "Documentos compartidos"
          }
        break;
      }
    }else{
      iProps = {
        endpointPlace: "Proyectos",
        projectId: "IdDelProyecto",
        urlProperty: "URLInternaDelEspacioDeTrabajoDelProyecto",
        documentPlace: "Documentos compartidos"
      }
    }
    let currentSiteUrl:Array<string> = document.URL.split("/");
    let isPDP:boolean = (currentSiteUrl.indexOf("Project%20Detail%20Pages") != -1)? true : false;
    if(isPDP){
      let id:string = document.URL.split('?')[1].split('&')[0].split('=')[1];
      let site:string = document.URL.toLowerCase().split('project')[0];
      let req:any = new XMLHttpRequest();
      req.open('GET', site + `_api/ProjectData/${iProps.endpointPlace}?$format=json&$filter=${iProps.projectId}%20eq%20guid%27${id}%27`);
      req.send('');
      req.onreadystatechange = () =>{
        if(req.readyState === 4){
          let parsed = JSON.parse(req.response);
          let projSite = iProps.urlProperty == "ProjectWorkspaceInternalUrl" ? parsed.value[0].ProjectWorkspaceInternalUrl : parsed.value[0].URLInternaDelEspacioDeTrabajoDelProyecto
          this.setState({
            url: projSite +`/${iProps.documentPlace}/Forms/AllItems.aspx`,
            loading: false,
            isPDP: true
          })
        }
      }
    }
    else{
      this.setState({loading:false});
    }
  }
  public render(): React.ReactElement<{IProjectFilesViewerProps}> {
    if(this.state.loading){
      return(
        <div>
          <h1>Loading...</h1>
        </div>
      )
    }else{
      if(this.state.isPDP){
        return (
          <div>
            <iframe src={this.state.url} style={{width: "100%", height: "500px"}}></iframe>
          </div>
        );
      }else{
        return( 
          <div>
            <h1>No es PDP</h1>
          </div>
        );
      }
    }
  }
}
