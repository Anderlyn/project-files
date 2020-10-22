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
      isPDP: false,
    };
  }
  componentDidUpdate(){
    if(this.state.isPDP){
      let workSpace = document.querySelector("#s4-workspace") as HTMLElement;
      workSpace.style.overflow = "hidden";
      let frameElementListener = document.getElementById("viewer") as HTMLElement;
      frameElementListener.onload = function(){
        let frameElement = document.getElementById("viewer") as HTMLIFrameElement;
        let header = frameElement.contentWindow.document.getElementsByClassName("ms-DetailsList-headerWrapper")[0] as HTMLElement;
        header.style.marginTop = "15px";
        workSpace.style.overflow = "auto";
        document.querySelector("#s4-workspace").scrollTop = 0;
        document.querySelector("#s4-workspace").scrollTop = 0;
        setTimeout(function(){
          document.getElementById("viewer").style.display = "initial"
        },1000);
      }
    }
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
    let currentSiteUrl:Array<string> = document.URL.toLowerCase().split("/");
    let isPDP:boolean = (currentSiteUrl.indexOf("project%20detail%20pages") != -1) ? true : false;
    if(isPDP){
      let id:string = document.URL.split('?')[1].split('&')[0].split('=')[1];
      let site:string = document.URL.toLowerCase().split('project')[0];
      let req:any = new XMLHttpRequest();
      req.open('GET', site + `_api/ProjectData/${iProps.endpointPlace}?$format=json&$filter=${iProps.projectId}%20eq%20guid%27${id}%27`);
      req.send('');
      req.onreadystatechange = () =>{
        let workSpace = document.querySelector("#s4-workspace") as HTMLElement;
        if(req.readyState === 4){
          let parsed = JSON.parse(req.response);
          let projSite = iProps.urlProperty == "ProjectWorkspaceInternalUrl" ? parsed.value[0].ProjectWorkspaceInternalUrl : parsed.value[0].URLInternaDelEspacioDeTrabajoDelProyecto
          this.setState({
            url: projSite +`/${iProps.documentPlace}/Forms/AllItems.aspx`,
            loading: false,
            isPDP: true
          });
          workSpace.scrollTop = 0;
        }
      }
    }else{
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
              <iframe id='viewer' src={this.state.url} style={{
              width: "100%", 
              display: "none",
              height: this.props.customHeight ? this.props.customHeight+"px" : "500px"
              }}></iframe>
          </div>
        );
      }else{
        return( 
          <div>
            <h1>Si estás leyendo esto, puede ser por las siguientes razones:</h1>
            <ul>
              <li>El webpart se utilizó en una página que no es PDP.</li>
              <li>El proyecto no tiene librería de documentos asociada.</li>
              <li>Hubo un error al traer la información del proyecto.</li>
            </ul>
          </div>
        );
      }
    }
  }
}
