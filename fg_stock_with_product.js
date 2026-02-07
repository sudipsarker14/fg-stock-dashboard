/** @odoo-module **/

import { registry } from "@web/core/registry";
import { Component, useState, onWillStart } from "@odoo/owl";
import { useService } from "@web/core/utils/hooks";
import { Many2One } from "@taps_hr/components/many2one/many2one";

class FGStockDashboard extends Component {
    static components = { Many2One };
    setup() {
        this.orm = useService("orm");
        this.userService = useService("user");

        this.state = useState({
            filters: {
                company_id: null,
                salesperson_id: null,
                team_id: null,
            },
            companies: [],
            salespersonOptions: [],
            teamOptions: [],
            selectedSalesperson: null,
            selectedTeam: null,
            data: {
                payment_terms: [],
                item_category: [],
                customer: [],
                buyer: [],
            },
            totals: {
                balance: 0,
                value: 0,
                balance_0_5: 0,
                value_0_5: 0,
                balance_30_plus: 0,
                value_30_plus: 0
            },
            table_totals: {
                payment_terms: {},
                item_category: {},
                customer: {},
                buyer: {},
            },
            showCustomer: false,
            showBuyer: false,
            loading: false,
            modal: {
                show: false,
                title: '',
                data: [],
                loading: false,
                context: {} // Store context for Excel download
            }
        });

        onWillStart(async () => {
            await this.loadCompanies();
            // Set default company to current user's company
            this.state.filters.company_id = this.userService.companyId;
            await this.loadDashboardData();
        });
    }

    async loadCompanies() {
        const companies = await this.orm.searchRead("res.company", [["id", "in", [1, 3]]], ["id", "name"]);
        this.state.companies = companies;

        // Load initial options for many2one fields
        this.state.salespersonOptions = await this.searchSalespersons("");
        this.state.teamOptions = await this.searchTeams("");
    }

    async loadDashboardData() {
        this.state.loading = true;
        try {
            const data = await this.orm.call(
                "fg.stock.dashboard",
                "get_dashboard_data",
                [this.state.filters]
            );
            this.state.data = data;

            // Initialize table totals
            this.state.table_totals = {
                payment_terms: this.calculateTableTotals(data.payment_terms),
                item_category: this.calculateTableTotals(data.item_category),
                customer: this.calculateTableTotals(data.customer),
                buyer: this.calculateTableTotals(data.buyer),
            };

            // Calculate overall summary totals
            let totalBalance = 0;
            let totalValue = 0;
            let total05Balance = 0;
            let total05Value = 0;
            let total30PlusBalance = 0;
            let total30PlusValue = 0;

            if (data.payment_terms && data.payment_terms.length > 0) {
                data.payment_terms.forEach(row => {
                    totalBalance += row.Total_Balance || 0;
                    totalValue += row.Total_Value || 0;
                    total05Balance += row['Balance_0-5'] || 0;
                    total05Value += row['Value_0-5'] || 0;
                    total30PlusBalance += row['Balance_30+'] || 0;
                    total30PlusValue += row['Value_30+'] || 0;
                });
            }
            this.state.totals.balance = totalBalance;
            this.state.totals.value = totalValue;
            this.state.totals.balance_0_5 = total05Balance;
            this.state.totals.value_0_5 = total05Value;
            this.state.totals.balance_30_plus = total30PlusBalance;
            this.state.totals.value_30_plus = total30PlusValue;
        } catch (error) {
            console.error("Error loading dashboard data:", error);
        } finally {
            this.state.loading = false;
        }
    }

    async onCompanyChange(ev) {
        this.state.filters.company_id = parseInt(ev.target.value);
        // Reload many2one options to sync with the new company
        this.state.salespersonOptions = await this.searchSalespersons("");
        this.state.teamOptions = await this.searchTeams("");
        await this.loadDashboardData();
    }

    async searchSalespersons(searchTerm) {
        const domain = searchTerm ? [['name', 'ilike', searchTerm], ['sale_team_id', '!=', false]] : [['sale_team_id', '!=', false]];
        const users = await this.orm.searchRead(
            "res.users",
            domain,
            ["id", "partner_id"],
            { limit: 0 }
        );
        const partnerIds = users.map(u => u.partner_id[0]);
        const partners = await this.orm.searchRead(
            "res.partner",
            [['id', 'in', partnerIds]],
            ["id", "name"]
        );
        const partnerMap = Object.fromEntries(partners.map(p => [p.id, p.name]));
        return users.map(u => ({ id: u.id, name: partnerMap[u.partner_id[0]] || 'Unknown' }));
    }

    async searchTeams(searchTerm) {
        const companyId = this.state.filters.company_id || false;
        let domain = [
            '&',
            ['is_sales', '=', true],
            '|',
            ['company_id', '=', false],
            ['company_id', '=', companyId]
        ];

        if (searchTerm) {
            domain = ['&', ['name', 'ilike', searchTerm], domain];
        }

        const teams = await this.orm.searchRead(
            "crm.team",
            domain,
            ["id", "name"],
            { limit: 0 }
        );
        return teams.map(t => ({ id: t.id, name: t.name }));
    }

    async onSalespersonSelect(record) {
        this.state.filters.salesperson_id = record ? record.id : null;
        this.state.selectedSalesperson = record;
        await this.loadDashboardData();
    }

    async onTeamSelect(record) {
        this.state.filters.team_id = record ? record.id : null;
        this.state.selectedTeam = record;
        await this.loadDashboardData();
    }

    calculateTableTotals(rows) {
        const totals = {
            'Balance_0-5': 0, 'Value_0-5': 0,
            'Balance_6-10': 0, 'Value_6-10': 0,
            'Balance_11-15': 0, 'Value_11-15': 0,
            'Balance_16-20': 0, 'Value_16-20': 0,
            'Balance_21-25': 0, 'Value_21-25': 0,
            'Balance_26-30': 0, 'Value_26-30': 0,
            'Balance_30+': 0, 'Value_30+': 0,
            'Total_Balance': 0, 'Total_Value': 0
        };

        if (rows && rows.length > 0) {
            rows.forEach(row => {
                Object.keys(totals).forEach(key => {
                    totals[key] += row[key] || 0;
                });
            });
        }
        return totals;
    }

    formatValue(value, decimals = 0) {
        if (!value && value !== 0) return '0 K';
        const formatted = (value / 1000).toLocaleString(undefined, {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        });
        return `${formatted} K`;
    }

    async onCellClick(row, field, groupType, bucketLabel) {
        if (row['Total'] === 'Total' || row['Buyer'] === 'Total' || row['Customer'] === 'Total' || row['Item'] === 'Total' || row['Payment Terms'] === 'Total') {
            // Logic to handle "Total" row clicks if needed, but the current code handles it
        }

        const groupField = groupType === 'payment_terms' ? 'Payment Terms' :
            groupType === 'item_category' ? 'Item' :
                groupType === 'customer' ? 'Customer' : 'Buyer';

        const groupValue = row[groupField] || 'Total';

        this.state.modal.show = true;
        this.state.modal.loading = true;
        this.state.modal.title = `Breakdown for ${groupField}: ${groupValue} (${bucketLabel})`;
        this.state.modal.context = { field, groupType, groupValue, bucketLabel };

        try {
            const data = await this.orm.call(
                "fg.stock.dashboard",
                "get_drilldown_data",
                [this.state.filters, groupType, groupValue, field]
            );
            this.state.modal.data = data;
        } catch (error) {
            console.error("Error fetching drilldown data:", error);
        } finally {
            this.state.modal.loading = false;
        }
    }

    get modalTotals() {
        const data = this.state.modal.data || [];
        return data.reduce((acc, item) => {
            acc.qty += (item.Qty || 0);
            acc.value += (item.Value || 0);
            return acc;
        }, { qty: 0, value: 0 });
    }

    async downloadDrilldownExcel() {
        const { field, groupType, groupValue, bucketLabel } = this.state.modal.context;
        try {
            const result = await this.orm.call(
                "fg.stock.dashboard",
                "download_drilldown_excel",
                [this.state.filters, groupType, groupValue, field]
            );

            const link = document.createElement('a');
            link.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + result.file_data;
            link.download = result.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } catch (error) {
            console.error("Error downloading drilldown Excel:", error);
        }
    }

    closeModal() {
        this.state.modal.show = false;
        this.state.modal.data = [];
    }

    async downloadExcel() {
        try {
            const result = await this.orm.call(
                "fg.stock.dashboard",
                "download_excel",
                [this.state.filters]
            );

            const link = document.createElement('a');
            link.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + result.file_data;
            link.download = result.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } catch (error) {
            console.error("Error downloading Excel:", error);
            alert("Failed to download Excel file. Please try again.");
        }
    }
}

FGStockDashboard.template = "fg_stock_dashboard_template";

registry.category("actions").add("action_fg_stock_dashboard", FGStockDashboard);
