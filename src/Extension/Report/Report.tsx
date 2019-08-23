import "./Report.scss";
import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IProjectPageService, getClient, IHostPageLayoutService } from "azure-devops-extension-api";
import { GitRestClient, VersionControlRecursionType } from "azure-devops-extension-api/Git";
import { ReleaseRestClient, ReleaseTaskAttachment } from "azure-devops-extension-api/Release";
import { Header, TitleSize } from "azure-devops-ui/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/HeaderCommandBar";
import { Page } from "azure-devops-ui/Page";
import { Spinner, SpinnerSize } from "azure-devops-ui/Spinner";
import { Tab, TabBar, TabSize } from "azure-devops-ui/Tabs";
import { showRootComponent } from "../../Common";

interface ItemReport {
    path?: string;
    comment?: string;
    url?: string;
    commitId?: string;
    content?: string;
    name?: string;
}

interface IReportContentState {
    selectedTabId: string;
    headerDescription?: string;
    useLargeTitle?: boolean;
    useCompactPivots?: boolean;
    items?: ItemReport[];
    nameEnvironment?: string;
    message?: string;
}

class ReportContent extends React.Component<{}, IReportContentState> {

    constructor(props: {}) {
        super(props);

        this.state = {
            selectedTabId: "0",
            items: [],
            message: "Loading ....."
        };
    }

    public componentDidMount() {

        SDK.register("registeredEnvironmentObject", () => {
            return {
                isInvisible: (tabContext: any) => {
                    if (tabContext && tabContext.releaseEnvironment && tabContext.releaseEnvironment.deployPhasesSnapshot) {
                        let phases = tabContext.releaseEnvironment.deployPhasesSnapshot;
                        let taskIds = [];
                        for (let phase of phases)
                            for (let task of phase.workflowTasks)
                                taskIds.push(task.taskId);
                        return taskIds.indexOf('13008e5e-9195-4043-9baf-c5d70da839e3') < 0;
                    }

                    return true;
                }
            }
        });

        SDK.init();
        this.initializeState();
    }

    private async initializeState(): Promise<void> {
        try {
            await SDK.ready();
            const configuration = await SDK.getConfiguration();
            const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
            const project = await projectService.getProject();

            if (configuration && configuration.releaseEnvironment && project) {
                const release = configuration.releaseEnvironment;
                const nameEnvironment = release.name;
                const RELEASE_RELEASEID = release.releaseId;
                const RELEASE_ENVIRONMENTID = release.id;
                const RELEASE_ATTEMPTNUMBER = release.deploySteps.length;
                const deploySteps: any[] = release.deploySteps;
                const deployStep = deploySteps[0];
                const releaseDeployPhases: any[] = deployStep.releaseDeployPhases;
                const runPlanIds = releaseDeployPhases.map(releaseDeployPhase => releaseDeployPhase.runPlanId);
                this.setState({ nameEnvironment: nameEnvironment });
                const releaseClient = getClient(ReleaseRestClient);

                const items: ItemReport[] = [];
                for (const planId of runPlanIds) {
                    const releaseTaskAttachments: ReleaseTaskAttachment[] = await releaseClient.getReleaseTaskAttachments(project.id, RELEASE_RELEASEID, RELEASE_ENVIRONMENTID, RELEASE_ATTEMPTNUMBER, planId, 'publish-report');
                    releaseTaskAttachments.forEach(attachment => {
                        const item: ItemReport = {
                            url: attachment.timelineId,
                            commitId: attachment.recordId,
                            path: planId,
                            comment: attachment.name.includes('-+-') ? attachment.name.split('-+-')[0]: attachment.name,
                            name: attachment.name
                        }
                        items.push(item);
                    });
                }

                for (const item of items) {
                    const content: ArrayBuffer = await releaseClient.getReleaseTaskAttachmentContent(project.id, RELEASE_RELEASEID, RELEASE_ENVIRONMENTID, RELEASE_ATTEMPTNUMBER, item.path as string, item.url as string, item.commitId as string, 'publish-report', item.name as string);
                    const dataView = new DataView(content);
                    const decoder = new TextDecoder("utf-8");
                    const decodedString = decoder.decode(dataView);
                    item.content = decodedString;
                }

                if (items.length > 0) {
                    this.setState({ selectedTabId: items[0].commitId as string });
                    this.setState({ items: items });
                } else {
                    await this.validateReportInRepository();
                }
            }

        } catch (error) {
            console.error("Error : ", error);
            this.setState({ message: " ¡ The report was not found !" });
        }

    }

    public render(): JSX.Element {

        const { selectedTabId, headerDescription, useCompactPivots, useLargeTitle, items, nameEnvironment, message } = this.state;

        if (items && items.length > 0) {

            const listItems = items.map((item) => <Tab name={item.comment} id={item.commitId as string} key={item.commitId} />);

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

                        {listItems}

                    </TabBar>

                    <div className="page-content">
                        <br></br>
                        {this.getPageContent()}
                    </div>
                </Page>
            );

        } else {

            return (
                <Page className="sample-hub flex-grow">
                    <div className="page-content">
                        <br /><br />
                        <Spinner className="page-content" size={SpinnerSize.large} label={message} />
                    </div>
                </Page>
            );
        }


    }

    private onSelectedTabChanged = (newTabId: string) => {
        this.setState({
            selectedTabId: newTabId
        })
    }

    private createMarkup(content: string) {
        return { __html: content };
    }

    private getPageContent() {
        const { selectedTabId, items } = this.state;
        if (items) {
            const item = items.find(x => x.commitId == selectedTabId);
            if (item) {
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
        const dialogService = await SDK.getService<IHostPageLayoutService>(CommonServiceIds.HostPageLayoutService);
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

    private async validateReportInRepository(): Promise<void> {
        try {

            const REPOSITORY_NAME: string = "DevOps_Vault_Reports_Extension";
            const configuration = await SDK.getConfiguration();
            const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
            const project = await projectService.getProject();

            if (configuration && configuration.releaseEnvironment && project) {

                const release = configuration.releaseEnvironment;
                const nameEnvironment = release.name;
                const RELEASE_DEFINITIONID = release.releaseDefinition.id;
                const RELEASE_RELEASEID = release.releaseId;
                const RELEASE_DEFINITIONENVIRONMENTID = release.definitionEnvironmentId;
                const RELEASE_ATTEMPTNUMBER = release.deploySteps.length;
                const path = `/${RELEASE_DEFINITIONID}/${RELEASE_RELEASEID}/${RELEASE_DEFINITIONENVIRONMENTID}/${RELEASE_ATTEMPTNUMBER}/`;

                this.setState({ nameEnvironment: nameEnvironment });

                const client = getClient(GitRestClient);
                let items = await client.getItems(REPOSITORY_NAME, project.id, path, VersionControlRecursionType.Full, true, true).then(items => {

                    return items.filter(item => !item.isFolder).map(elem => {

                        let item: ItemReport = {
                            url: elem.url,
                            commitId: elem.latestProcessedChange.commitId,
                            path: elem.path,
                            comment: elem.latestProcessedChange.comment
                        }

                        return item;
                    });

                });

                for (let index = 0; index < items.length; index++) {
                    items[index].content = await client.getItemText(REPOSITORY_NAME, items[index].path as string, project.id);
                }

                if (items.length > 0) {
                    this.setState({ selectedTabId: items[0].commitId as string });
                    this.setState({ items: items });
                } else {
                    await this.validateReportOld();
                }

            }

        } catch (err) {
            console.error("Error in validateReportInRepository : ", err);
            this.setState({ message: " ¡ The report was not found !" });
        }
    }

    private async validateReportOld(): Promise<void> {

        try {
            const idTask = "13008e5e-9195-4043-9baf-c5d70da839e3";
            const configuration = await SDK.getConfiguration();
            const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
            const project = await projectService.getProject();

            if (configuration && configuration.releaseEnvironment && project) {

                const release = configuration.releaseEnvironment;
                const nameEnvironment = release.name;
                const RELEASE_DEFINITIONID = release.releaseDefinition.id;
                const RELEASE_RELEASEID = release.releaseId;
                const path = RELEASE_DEFINITIONID + "/" + RELEASE_RELEASEID + ".html";

                this.setState({ nameEnvironment: nameEnvironment });

                const REPOSITORY_NAME = release.deployPhasesSnapshot.find((x: any) => {
                    return x.workflowTasks.find((y: any) => y.taskId === idTask)
                }).workflowTasks.find((x: any) => x.taskId === idTask).inputs.nameRepository;

                if (REPOSITORY_NAME) {

                    const client = getClient(GitRestClient);

                    const item: ItemReport = {
                        url: "",
                        commitId: Math.random().toString(36).substring(7),
                        path: path,
                        comment: "Report"
                    }

                    item.content = await client.getItemText(REPOSITORY_NAME, item.path as string, project.id);

                    const items = [item];

                    if (item.content && item.content.length > 0) {
                        this.setState({ selectedTabId: items[0].commitId as string });
                        this.setState({ items: items });
                    } else {
                        this.setState({ message: " ¡ The report was not found !" });
                    }

                } else {
                    throw new Error('is valid if it belongs to an old repository but the repository was not found');
                }
            }

        } catch (error) {
            console.error("Error in validateReportOld : ", error);
            this.setState({ message: " ¡ The report was not found !" });
        }
    }
}

showRootComponent(<ReportContent />);
