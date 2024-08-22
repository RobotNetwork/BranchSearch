# Branch Search Web Part

### Summary

The Branch Search Web Part allows users to search and display branch information within a SharePoint site. It leverages the SharePoint Framework (SPFx) and integrates seamlessly with Microsoft 365.



### Features

- Search for branch information by branch number
- Display branch details including address, phone, fax, manager, and more

### Applies to

- [SharePoint Framework](https://aka.ms/spfx) ![version](https://img.shields.io/badge/version-1.19.0-green.svg)

### Prerequisites

- SharePoint Online (SPO)
- Permissions to add and configure web parts

### Getting Started

1. Clone the repository
  ```
  git clone https://github.com/your-repo/branch-search-webpart.git
  ```
2. Navigate to the project directory
  ```
  cd branch-search-webpart
  ```
3. Install dependencies
  ```
  npm install
  ```
4. Build the solution
  ```
  gulp build
  ```
5. Bundle the solution
  ```
  gulp bundle --ship
  ```
6. Package the solution
```
gulp package-solution --ship
```

### Deployment
1. Upload the .sppkg file from the sharepoint/solution folder to your SharePoint App Catalog.
2. Add the web part to your SharePoint site by navigating to the site or site collection, Settings -> Site Contents -> Add an App
  2a. Only necessary if the option to make available for all sites was unchecked. If installed for all sites, skip this step
3. Add the web part to any desired section of SPO.