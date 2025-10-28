# Company Anniversaries and Birthdays Web Part - Deployment Guide

## Overview
This SharePoint Framework (SPFx) web part displays employee anniversaries and birthdays in a visually appealing way with multiple display options.

## Features
- **Automatic List Creation**: On first use, the web part automatically creates an "Employee Celebrations" list with required columns
- **Multiple Display Modes**: Grid, List, and Carousel views
- **Flexible Filtering**: Show events for today, this week, this month, next month, or all upcoming
- **Visual Appeal**: Beautiful gradient cards with different colors for birthdays (pink) and anniversaries (blue)
- **Animated Icons**: Birthday cake and celebration emojis with bounce animations
- **Smart Date Calculation**: Automatically calculates next occurrence of birthdays and anniversaries
- **Responsive Design**: Works on desktop, tablet, and mobile devices

## Prerequisites
- Node.js (v16 or v18 recommended)
- SharePoint Online tenant
- Admin access to SharePoint App Catalog

## Build and Package

### 1. Build the Solution
```bash
npm install
npm run build
```

### 2. Bundle for Production
```bash
gulp bundle --ship
```

### 3. Package the Solution
```bash
gulp package-solution --ship
```

This will create a `.sppkg` file in the `sharepoint/solution` folder.

## Deploy to SharePoint

### 1. Upload to App Catalog
1. Go to your SharePoint App Catalog site
2. Navigate to **Apps for SharePoint**
3. Upload the `.sppkg` file from `sharepoint/solution` folder
4. Check "Make this solution available to all sites in the organization" if you want
5. Click **Deploy**

### 2. Add the Web Part to a Page
1. Navigate to any modern SharePoint page
2. Edit the page
3. Click the + icon to add a web part
4. Search for "Company Anniversaries and Birthdays"
5. Add it to the page

### 3. Configure the Web Part
The web part will automatically create an "Employee Celebrations" list on first use. You can configure:
- **Select Employee List**: Choose which SharePoint list to use (defaults to "Employee Celebrations")
- **Display Mode**: Grid, List, or Carousel view
- **Show Events**: Filter by Today, This Week, This Month, Next Month, or All Upcoming
- **Show Images**: Toggle birthday and anniversary icons on/off

## SharePoint List Structure

The web part expects a SharePoint list with the following columns:

| Column Name | Type | Description |
|------------|------|-------------|
| Title | Single line of text | Employee name (created by default) |
| HireDate | Date | Employee hire date for anniversaries |
| Birthday | Date | Employee birth date |

### Adding Employee Data
1. Go to the "Employee Celebrations" list
2. Click "New" to add an employee
3. Fill in:
   - **Title**: Employee's full name
   - **HireDate**: Date they were hired
   - **Birthday**: Their birth date (year doesn't matter for display)

## How It Works

### Date Calculation
- The web part calculates the **next occurrence** of each birthday and anniversary
- For birthdays, it uses the month and day from the Birthday column
- For anniversaries, it uses the month and day from the HireDate column
- It calculates years of service based on the hire date

### Display Features
- **Birthday Cards**: Pink gradient with 🎂 icon
- **Anniversary Cards**: Blue gradient with 🎉 icon and years of service
- **Carousel Mode**: Auto-rotates through events every 5 seconds
- **Hover Effects**: Cards lift and show shadow on hover

## Customization

### Changing Colors
Edit `src/webparts/companyAnniversariesBirthdays/components/CompanyAnniversariesBirthdays.module.scss`:

```scss
&.birthday {
  background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
}

&.anniversary {
  background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
}
```

### Changing Carousel Speed
Edit line 84 in `src/webparts/companyAnniversariesBirthdays/components/CompanyAnniversariesBirthdays.tsx`:

```typescript
}, 5000); // Change from 5000ms (5 seconds) to desired milliseconds
```

## Troubleshooting

### List Not Found Error
- Verify the list name matches in the property pane
- Ensure the list has the required HireDate and Birthday columns
- Check SharePoint permissions

### No Events Showing
- Verify employees have been added to the list
- Check that HireDate and Birthday columns have valid dates
- Try changing the filter mode in the property pane

### Build Errors
- Run `npm install` to ensure all dependencies are installed
- Clear the cache: `gulp clean`
- Rebuild: `npm run build`

## Support and Maintenance

### Updating the Web Part
1. Make changes to the source code
2. Increment the version in `package-solution.json`
3. Rebuild and repackage
4. Upload the new `.sppkg` to App Catalog
5. SharePoint will prompt to update

### Testing Locally
```bash
gulp serve
```
This will start the local workbench for testing.

## API Permissions
This web part uses standard SharePoint REST APIs and requires:
- Read access to SharePoint lists
- Write access to create the list (first time only)

No additional API permissions are required.

## Browser Support
- Microsoft Edge (recommended)
- Google Chrome
- Mozilla Firefox
- Safari

## Version History
- **1.0.0**: Initial release with grid, list, and carousel views
