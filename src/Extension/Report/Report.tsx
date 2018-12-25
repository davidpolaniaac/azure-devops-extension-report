import "./Report.scss";

import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IHostDialogService, IProjectPageService, getClient} from "azure-devops-extension-api";
import { GitRestClient, VersionControlRecursionType } from "azure-devops-extension-api/Git";
import { Header, TitleSize } from "azure-devops-ui/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/HeaderCommandBar";
import { Page } from "azure-devops-ui/Page";
import { Tab, TabBar, TabSize } from "azure-devops-ui/Tabs";
import { showRootComponent } from "../../Common";

interface ItemReport {
    path?:string;
    comment?:string;
    url?:string;
    commitId?: string;
    content?: string;
}

interface IReportContentState {
    selectedTabId: string;
    headerDescription?: string;
    useLargeTitle?: boolean;
    useCompactPivots?: boolean;
    items?: ItemReport[];
    nameEnvironment?: string;

}

class ReportContent extends React.Component<{}, IReportContentState> {

    constructor(props: {}) {
        super(props);

        this.state = {
            selectedTabId: "0", 
            items: []
        };
    }

    public async componentDidMount() {
        await SDK.init();
        await SDK.ready();
        await this.initializeState();
    }

    private async initializeState(): Promise<void> {

        const REPOSITORY_NAME: string = "DevOps_Vault_Reports_Extension";
        const configuration = await SDK.getConfiguration();
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();

        if(configuration && project ){

            const release = configuration.releaseEnvironment;
            const nameEnvironment =  release.name;
            const RELEASE_DEFINITIONID = release.releaseDefinition.id;
            const RELEASE_RELEASEID = release.releaseId;
            const RELEASE_DEFINITIONENVIRONMENTID = release.definitionEnvironmentId;
            const RELEASE_ATTEMPTNUMBER = release.deploySteps.length; 
            const path = `/${RELEASE_DEFINITIONID}/${RELEASE_RELEASEID}/${RELEASE_DEFINITIONENVIRONMENTID}/${RELEASE_ATTEMPTNUMBER}/`;
    
            this.setState({ nameEnvironment: nameEnvironment });

            const client = getClient(GitRestClient);
            let items = await client.getItems(REPOSITORY_NAME,project.id,path,VersionControlRecursionType.Full).then( items => {
                
                return items.filter(item => !item.isFolder).map(elem => {
    
                    let item:ItemReport = {
                        url: elem.url,
                        commitId: elem.commitId,
                        path: elem.path,
                      } ;
    
                    return item;
                });
    
            });
    
            for (let index = 0; index < items.length; index++) {
                items[index].comment = await client.getCommit(items[index].commitId as string,REPOSITORY_NAME,project.id).then( commit => commit.comment);
                items[index].content  = await client.getItemText(REPOSITORY_NAME,items[index].path as string , project.id);
            }
            
            this.setState({ items: items });

        }
    }

    public render(): JSX.Element {

        const { selectedTabId, headerDescription, useCompactPivots, useLargeTitle, items , nameEnvironment} = this.state;
        
        if(items){
            const listItems = items.map((item) => <Tab name={item.comment} id={item.commitId as string}  key={item.commitId}/> );

            return (
                <Page className="sample-hub flex-grow">
    
                    <Header title={nameEnvironment}
                        commandBarItems={this.getCommandBarItems()}
                        description={headerDescription}
                        titleSize={useLargeTitle ? TitleSize.Large : TitleSize.Medium} />
    
                    <TabBar
                        onSelectedTabChanged={this.onSelectedTabChanged}
                        selectedTabId={selectedTabId}
                        tabSize={useCompactPivots ? TabSize.Compact : TabSize.Tall}>
    
                        { listItems }
    
                    </TabBar>
    
                    <div className="page-content">
                    <br></br>
                        { this.getPageContent() }
                    </div>
                </Page>
            );

        }else{

            return ( 
                <Page className="sample-hub flex-grow"> </Page>
            );
        }

        
    }

    private onSelectedTabChanged = (newTabId: string) => {
        this.setState({
            selectedTabId: newTabId
        })
    }

    private createMarkup(content : string ) {
        return {__html: content };
    }

    private getPageContent() {
        const { selectedTabId, items } = this.state;
        if(items){
            const item = items.find( x => x.commitId == selectedTabId);
            if(item){
                return <div dangerouslySetInnerHTML={this.createMarkup(item.content as string)} />;
            }
        }
    }

    private getCommandBarItems(): IHeaderCommandBarItem[] {
        return [
            {
              id: "download",
              text: "Download",
              onActivate: () => { this.onCustomPromptClick() },
              iconProps: {
                iconName: 'Download'
              },
              isPrimary: true,
              tooltipProps: {
                text: "Download a report"
              }
            }
        ];
    }

    private async onCustomPromptClick(): Promise<void> {
        const dialogService = await SDK.getService<IHostDialogService>(CommonServiceIds.HostDialogService);
        dialogService.openCustomDialog<boolean | undefined>(SDK.getExtensionContext().id + ".panel-content", {
            title: "Soon",
            configuration: {
                message: "Soon you will be able to download the reports",
                initialValue: this.state.useCompactPivots
            },
            onClose: (result) => {
                if (result !== undefined) {
                    this.setState({ useCompactPivots: result });
                }
            }
        });
    }
}

showRootComponent(<ReportContent />);