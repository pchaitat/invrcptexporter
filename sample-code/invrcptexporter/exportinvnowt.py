#!/usr/bin/env python3

from invrcpt.exporter import Ods, INV_NOWT_FORM

list_rows = (2, 3, 4)
ods = Ods()
ods.export_multiple_common_form_to_pdf(
  form=INV_NOWT_FORM,
  list_rows=list_rows,
  dest='/home/invrcptexporter-dev/tmp/')

