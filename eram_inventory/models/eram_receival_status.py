# -*- coding: utf-8 -*-
from odoo import fields, models


class EramReceivalStatus(models.Model):
    _name = "eram.receival.status"
    _description = "Eram Receival Status"

    name = fields.Char("Name")
