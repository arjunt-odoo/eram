# -*- coding: utf-8 -*-
from odoo import fields, models


class EramReceivalStatus(models.Model):
    _name = "eram.receival.status"

    name = fields.Char("Name")
