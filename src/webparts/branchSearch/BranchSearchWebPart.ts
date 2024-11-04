import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
} from "@microsoft/sp-property-pane";

import styles from "./BranchSearchWebPart.module.scss";

export interface IGetSpListItemsWebPartProps {
    description: string;
}

interface GoogleSheetLists {
    value: GoogleSheetList[];
}

interface GoogleSheetList {
    locationCode: number;
    locationName: string;
    brand: string;
    division: string;
    region: string;
    manager: string;
    addressLine1: string;
    addressLine2: string;
    city: string;
    state: string;
    zipCode: string;
    phone: string;
    fax: string;
    emergency: string;
    branchManagerEmail: string;
    branchManagerPhoneNumber: string;
    rvp: string;
    rvpPhoneNumber: string;
    president: string;
    locationAttributes: string;
    plumbing: string;
    waterworks: string;
    hvac: string;
    bK: string;
    headquarters: string;
    gmtOffset: string;
    daylightSavingsTime: string;
    customerFacing: string;
}

export default class GetSpListItemsWebPart extends BaseClientSideWebPart<IGetSpListItemsWebPartProps> {
    private _getListData(input?: string): Promise<GoogleSheetList[]> {
        
        const sheetUrl = `https://docs.google.com/spreadsheets/d/${this.properties.description}/gviz/tq?tqx=out:json`;
        return fetch(sheetUrl)
            .then((response) => response.text()) // Read as plain text
            .then((text) => {
                // Remove the unnecessary prefix "google.visualization.Query.setResponse("
                const json = JSON.parse(text.substring(47, text.length - 2));

                const rows = json.table.rows.map(
                    (row: { c: { v: string | number }[] }) => ({
                        locationCode: row.c[0]?.v || null,
                        locationName: row.c[1]?.v || "",
                        brand: row.c[2]?.v || "",
                        division: row.c[3]?.v || "",
                        region: row.c[4]?.v || "",
                        manager: row.c[5]?.v || "",
                        addressLine1: row.c[6]?.v || "",
                        addressLine2: row.c[7]?.v || "",
                        city: row.c[8]?.v || "",
                        state: row.c[9]?.v || "",
                        zipCode: row.c[10]?.v || null,
                        phone: row.c[11]?.v || "",
                        fax: row.c[12]?.v || "",
                        emergency: row.c[13]?.v || "",
                        branchManagerEmail: row.c[14]?.v || "",
                        branchManagerPhoneNumber: row.c[15]?.v || "",
                        rvp: row.c[16]?.v || "",
                        rvpPhoneNumber: row.c[17]?.v || "",
                        president: row.c[18]?.v || "",
                        locationAttributes: row.c[19]?.v || "",
                        plumbing: row.c[20]?.v || "",
                        waterworks: row.c[21]?.v || "",
                        hvac: row.c[22]?.v || "",
                        bK: row.c[23]?.v || "",
                        headquarters: row.c[24]?.v || "",
                        gmtOffset: row.c[25]?.v || null,
                        daylightSavingsTime: row.c[26]?.v || "",
                        customerFacing: row.c[27]?.v || "",
                    })
                );

                if (!input) return rows;

                return rows.filter(
                    (item: {
                        locationCode: number;
                        city: string;
                        state: string;
                    }) =>
                        item.locationCode === Number(input) ||
                        item.city.toLowerCase() === input.toLowerCase() ||
                        item.state.toLowerCase() === input.toLowerCase()
                );
            })
            .catch((error) =>
                console.error("Error fetching data from Google Sheets: ", error)
            );
    }

    private _getSearchField(): HTMLInputElement {
        const ele = this.domElement.querySelector(
            "#branchSearchInput"
        ) as HTMLInputElement;
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
                            this._renderBranchData({ value: data });
                        })
                        .catch((error) => {
                            console.error(error);
                        });
                });
            });
        }, 1000);
    }

    private _renderList(data: GoogleSheetLists): void {
        const html = data.value
            .map(
                (item) =>
                    `
			<div class="${styles.branchCard}">
				<h2 class="${styles.title}">${item.locationCode}</h2>
				<div>${item.addressLine1} ${item.addressLine2 ? item.addressLine2 : ""} ${item.city}, ${
                        item.state
                    } ${item.zipCode}</div>
				<div>${item.phone}</div>
			</div>
		`
            )
            .join("");
        this._setHtml(html);
        this._setCardEventListeners();
    }

    private _renderBranchData(data: GoogleSheetLists): void {
        data.value.map((item) => {
            this._getSearchField().value = item.locationCode.toString();
        });

        const html = data.value
            .map(
                (item) =>
                    `
                <div class="${
                    styles.moreInfo
                }"><a href="https://docs.google.com/spreadsheets/d/${this.properties.description}" target="_blank">Click here</a>&nbsp;for all branch info!</div>
				<div class="${styles.branchData}">
					<h2 class="${styles.title}">${item.locationCode}</h2>
					<div><strong>Address:</strong> ${item.addressLine1} ${
                        item.addressLine2 ? item.addressLine2 : ""
                    } ${item.city}, ${item.state} ${item.zipCode}</div>
					<div><strong>Region:</strong> ${item.region}</div>
					<div><strong>Phone:</strong> ${item.phone}</div>
					<div><strong>Fax:</strong> ${item.fax}</div>
					<div><strong>Manager:</strong> ${item.manager} </div>
					<div><strong>RVP:</strong> ${item.rvp}</div>
					<div><strong>Hours:</strong> ${item.locationAttributes}</div>
					<div><strong>Emergency:</strong> ${item.emergency}</div>
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
                    .then((data) => this._renderList({ value: data }))
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
            .then((data) => this._renderList({ value: data }))
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
                        description: "Settings",
                    },
                    groups: [
                        {
                            groupFields: [
                                PropertyPaneTextField("description", {
                                    label: "Google Sheet ID",
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    }
}
