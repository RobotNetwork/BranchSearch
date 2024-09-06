import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
} from "@microsoft/sp-property-pane";

import styles from "./BranchSearchWebPart.module.scss";
// import * as strings from "BranchSearchWebPartStrings";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
// import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

export interface IGetSpListItemsWebPartProps {
    description: string;
}

export interface ISPLists {
    value: ISPList[];
}

// export interface ISPList {
//     Title: string; // branch number
//     field_3: string; // division
//     field_4: string; // region
//     field_5: string; // manager
//     field_6: string; // address
//     field_7: string; // city
//     field_8: string; // state
//     field_9: string; // zip code
//     field_10: string; // phone number
//     field_11: string; // fax number
//     field_12: string; // emergency contact
//     field_13: { 0: { EMail: string } }; // manager email
//     field_14: string; // manager contact number
//     field_15: string; // rvp
//     field_16: string; // rvp cell
//     field_17: string; // operating hours
// }

export interface ISPList {
    Title: string; // branch number
    field_3: string; // division
    field_4: string; // region
    field_5: string; // manager
    field_6: string; // address
    field_7: string; // suite
    field_8: string; // city
    field_9: string; // state
    field_10: string; // zip code
    field_11: string; // phone number
    field_12: string; // fax number
    field_13: string; // emergency contact
    field_14: { 0: { EMail: string } }; // manager email
    field_15: string; // manager contact number
    field_16: string; // rvp
    field_17: string; // rvp cell
    field_19: string; // operating hours
}

export default class GetSpListItemsWebPart extends BaseClientSideWebPart<IGetSpListItemsWebPartProps> {
    
    private _getListData(input?: string): Promise<ISPLists> {
        let requestUrl =
        "https://morscousa.sharepoint.com/sites/ReeceHub/_api/web/lists/GetByTitle('Locations')/Items?$select=Title,field_3,field_4,field_5,field_6,field_7,field_8,field_9,field_10,field_11,field_12,field_13,field_14/EMail,field_15,field_16,field_17,field_19&$expand=field_14";

        if (input) {
            const filterQuery = `&$filter=substringof('${input}', Title) or substringof('${input}', field_8) or substringof('${input}', field_9)`;
            requestUrl += filterQuery;
        }

        return this.context.spHttpClient
            .get(requestUrl, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            })
            .catch((error) => {
                console.error(error);
            });
    }

    private _getSearchField(): HTMLInputElement {
        const ele = this.domElement.querySelector("#branchSearchInput") as HTMLInputElement;
        return ele;
    }

    private _getListContainer(): Element | null {
        const listContainer = this.domElement.querySelector(
            `.${styles.spListcontainer}`
        );
        if (!listContainer) {
            console.error(
                `Element with ID '${styles.spListcontainer}' not found.`
            );
            return null;
        }
        return listContainer;
    }

    private _getBaseHtml(): string {
        return `
			<div class="${styles.inputContainer}">
				<svg
					class="${styles.searchIcon}"
					xmlns="http://www.w3.org/2000/svg"
					width="36"
					height="36"
					viewBox="0 0 24 24"
					fill="none"
					stroke="currentColor"
					stroke-width="2"
					stroke-linecap="round"
					stroke-linejoin="round"
				>
					<circle cx="11" cy="11" r="8"></circle>
					<path d="m21 21-4.3-4.3"></path>
				</svg>
				<input
					class="${styles.inputStyle}"
					placeholder="Branch #, City, or State"
					type="search"
					id="branchSearchInput"
				/>
			</div>
			<div class="${styles.spListcontainer}"/>
    	`;
    }

    private _setHtml(html: string): void {
        const listContainer = this._getListContainer();
        if (!listContainer) return;

        listContainer.innerHTML = html;
    }

    private _setCardEventListeners(): void {
        setTimeout(() => {
            const cards = this.domElement.querySelectorAll(
                `.${styles.branchCard}`
            ) as NodeListOf<HTMLHeadingElement>;
            cards.forEach((card) => {
                card.addEventListener("click", () => {
                    const branchNumber = card.innerText.split("\n")[0];
                    this._getListData(branchNumber)
                        .then((data) => {
                            this._renderBranchData(data);
                        })
                        .catch((error) => {
                            console.error(error);
                        });
                });
            });
        }, 1000);
    }

    private _renderList(data: ISPLists): void {
        const html = data.value
            .map(
                (item) =>
                    `
			<div class="${styles.branchCard}">
				<h2 class="${styles.title}">${item.Title}</h2>
				<div>${item.field_6} ${item.field_7 ? item.field_7 : ""} ${item.field_8}, ${item.field_9} ${item.field_10}</div>
				<div>${item.field_11}</div>
			</div>
		`
            )
            .join("");
        this._setHtml(html);
        this._setCardEventListeners();
    }

    private _renderBranchData(data: ISPLists): void {

        data.value.map((item) => {
            this._getSearchField().value = item.Title;
        });

        const html = data.value
            .map(
                (item) => 
                    `
                <div class="${styles.moreInfo}"><a href="https://morscousa.sharepoint.com/sites/ReeceHub/Lists/Locations/AllItems.aspx?q=${this._getSearchField().value}" target="_blank">Click here</a>&nbsp;for more branch info!</div>
				<div class="${styles.branchData}">
					<h2 class="${styles.title}">${item.Title}</h2>
					<div><strong>Address:</strong> ${item.field_6} ${item.field_7 ? item.field_7 : ""} ${item.field_8}, ${
                        item.field_9
                    } ${item.field_10}</div>
					<div><strong>Region:</strong> ${item.field_4}</div>
					<div><strong>Phone:</strong> ${item.field_11}</div>
					<div><strong>Fax:</strong> ${item.field_12}</div>
					<div><strong>Manager:</strong> <a href="mailto:${item.field_14[0].EMail}"> ${
                        item.field_5
                    } </a> </div>
					<div><strong>RVP:</strong> ${item.field_16}</div>
					<div><strong>Hours:</strong> ${item.field_19}</div>
					<div><strong>Emergency:</strong> ${item.field_12}</div>
				</div>
			`
            )
            .join("");

        this._setHtml(html);
        this._setCardEventListeners();
    }

    private _setSearchEventListener(): void {
        const searchInput = this._getSearchField();
        if (searchInput) {
            searchInput.addEventListener("input", () => {
                const branchNumber = searchInput.value;
                this._getListData(branchNumber)
                    .then((data) => this._renderList(data))
                    .catch((error) => {
                        console.error(error);
                    });
            });
        }
    }

    public render(): void {
        this.domElement.innerHTML = this._getBaseHtml();

        this._setSearchEventListener();
        this._getListData()
            .then((data) => 
                this._renderList(data))
            .catch((error) => {
                console.error(error);
            });
        this._setCardEventListeners();
    }

    protected get dataVersion(): Version {
        return Version.parse("1.0.2.0");
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: "Basic description",
                    },
                    groups: [
                        {
                            groupName: "Group Name",
                            groupFields: [
                                PropertyPaneTextField("description", {
                                    label: "Description",
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    }
}
