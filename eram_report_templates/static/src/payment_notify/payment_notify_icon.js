/** @odoo-module **/
import { registry } from "@web/core/registry";
import { Component, useState, onWillStart } from "@odoo/owl";
import { useService } from "@web/core/utils/hooks";
import { Dropdown } from "@web/core/dropdown/dropdown";
import { useDropdownState } from "@web/core/dropdown/dropdown_hooks";
import { useDiscussSystray } from "@mail/utils/common/hooks";

export class PaymentAlertSystray extends Component {
    static components = { Dropdown };
    static template = "eram_report_templates.PaymentAlertSystray";

    setup() {
        this.dialogService = useService("dialog");
        this.actionService = useService("action");
        this.discussSystray = useDiscussSystray();
        this.orm = useService("orm");
        this.dropdown = useDropdownState();
        this.state = useState({
            alerts: [],
            hasOverduePayments: false
        });

        onWillStart(async () => {
            await this.loadOverdueInvoices();
        });
    }
    async loadOverdueInvoices() {
        await this.orm.searchRead(
            "account.move",
            [
                ['move_type', '=', 'out_invoice'],
                ['state', '=', 'posted'],
                ['invoice_date_due', '!=', false],
                ['amount_residual', '>', 0]
            ],
            ["invoice_date_due", "name", "amount_residual"]
        ).then((res) => {
            if (res.length) {
                const today = new Date();
                today.setHours(0, 0, 0, 0);

                const overdueRecords = res.filter(record => {
                    if (!record.invoice_date_due) return false;
                    const dueDate = new Date(record.invoice_date_due);
                    dueDate.setHours(0, 0, 0, 0);
                    return dueDate <= today;
                });

                this.state.alerts = overdueRecords;
                this.state.hasOverduePayments = overdueRecords.length > 0;
            }
        });
    }
    openInvoice(invoiceId) {
        this.dropdown.close();
        this.actionService.doAction({
            type: 'ir.actions.act_window',
            res_model: 'account.move',
            res_id: invoiceId,
            views: [[false, 'form']],
            target: 'current',
        });
    }
}

export const systrayItem = {
    Component: PaymentAlertSystray,
};

registry.category("systray").add("PaymentAlertSystray", systrayItem, { sequence: 1 });