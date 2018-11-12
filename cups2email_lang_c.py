#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
from cups2email_lang import cups2email

cups2email(
    filename=sys.argv[1],
    template='Petición documentación: Suministros >20 años_2avis'
)
