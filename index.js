(async function () {
    const formFields = {
        'sourceSheet': 'selectedWorksheet',
        'parameter': 'pickparam',
        'virtualTable': 'virtualWorksheet'
    };

    async function init() {
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        const worksheets = dashboard.worksheets;
        const isConfigured = tableau.extensions.settings.get('configured') === 'true';
        await populateWorksheetDropdown(formFields.sourceSheet);
        await populateWorksheetDropdown(formFields.virtualTable);
        await populateParameterDropdown();

        if (isConfigured) {
            const selectedWorksheetName = tableau.extensions.settings.get(formFields.sourceSheet);
            const worksheet = worksheets.find(sheet => sheet.name === selectedWorksheetName);

            worksheet.addEventListener(tableau.TableauEventType.FilterChanged, () => {
                runExtension(dashboard, worksheet);
            });

            runExtension(dashboard, worksheet);
        }
    }

    async function runExtension(dashboard, worksheet) {
        const selectedParameterName = tableau.extensions.settings.get(formFields.parameter);
        const virtualTableWorksheet = tableau.extensions.settings.get(formFields.virtualTable);
        const selectedParameter = await dashboard.findParameterAsync(selectedParameterName);
        const dataTableReader = await worksheet.getSummaryDataReaderAsync();
        const dataTable = await dataTableReader.getAllPagesAsync();
        await dataTableReader.releaseAsync();

        const data = [
            dataTable.columns.map(col => col.fieldName),
            dataTable.data.map(row => row.map(cell => cell.nativeValue))
        ];

        selectedParameter.changeValueAsync(JSON.stringify(data)).then(() => {
            dashboard.worksheets.forEach(sheet => {
                if (sheet.name === virtualTableWorksheet) {
                    sheet.getDataSourcesAsync().then(datasources => {
                        datasources[0].refreshAsync();
                    });
                }
            });
        });
    }

    async function populateWorksheetDropdown(elementId) {
        const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
        const selectedWorksheetName = tableau.extensions.settings.get(elementId);
        const options = worksheets.map(worksheet => {
            const isSelected = worksheet.name === selectedWorksheetName;
            return `<option value='${worksheet.name}' ${isSelected ? 'selected' : ''}>${worksheet.name}</option>`;
        }).join('');
        document.getElementById(elementId).innerHTML = options;
        document.getElementById('save-settings').disabled = false;
    }

    async function populateParameterDropdown() {
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        const selectedParameterName = tableau.extensions.settings.get(formFields.parameter);
        const params = await dashboard.getParametersAsync();
        const stringParams = params.filter(p => p.dataType === 'string');
        const options = stringParams.map(param => {
            const isSelected = param.name === selectedParameterName;
            return `<option value='${param.name}' ${isSelected ? 'selected' : ''}>${param.name}</option>`;
        }).join('');

        document.getElementById(formFields.parameter).innerHTML = options || "<option value='' disabled>No parameters found</option>";
        document.getElementById('save-settings').disabled = false;
    }

    function getInstanceConfigs(element) {
        const selectedWorksheet = element.querySelector(`#${formFields.sourceSheet}`).value;
        const selectedParam = element.querySelector(`#${formFields.parameter}`).value;
        const virtualTableWorksheet = element.querySelector(`#${formFields.virtualTable}`).value;

        return {
            selectedWorksheet,
            selectedParam,
            virtualTableWorksheet
        };
    }

    async function submit() {
        const instanceConfig = getInstanceConfigs(document.querySelector('.extension-container'));

        tableau.extensions.settings.set(formFields.parameter, instanceConfig.selectedParam);
        tableau.extensions.settings.set(formFields.sourceSheet, instanceConfig.selectedWorksheet);
        tableau.extensions.settings.set(formFields.virtualTable, instanceConfig.virtualTableWorksheet);
        tableau.extensions.settings.set('configured', 'true');

        await tableau.extensions.settings.saveAsync();
        init();
    }

    document.addEventListener('DOMContentLoaded', () => {
        tableau.extensions.initializeAsync().then(() => {
            init();

            document.getElementById('save-settings').addEventListener('click', function () {
                submit();
            });

            document.getElementById('initialize').addEventListener('click', function () {
                init();
            });
        }, (err) => {
            console.error("Initialization failed:", err);
        });
    });
})();