from datetime import datetime
from pathlib import Path

import pynetbox
import xlsxwriter

import config

if __name__ == "__main__":
    # excel report name
    current_dir = Path(__file__).parent.absolute()
    report_file = current_dir.joinpath(
        current_dir,
        f"{config.REPORT_NAME}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )

    # getting info from netbox
    nb = pynetbox.api(
        url=config.NETBOX_URL,
        token=config.NETBOX_AUTH_TOKEN,
        threading=True,
    )
    vlan_list = nb.ipam.vlans.all()

    # saving vlan list to excel file
    # column names
    header = [
        {"header": "VLAN_ID"},
        {"header": "Name"},
        {"header": "Status"},
        {"header": "Tags"},
    ]

    # formating data
    excel_data = []
    for vlan in vlan_list:
        str_tags = [str(tag) for tag in vlan.tags]
        excel_data.append(
            [
                str(vlan.vid),
                str(vlan.name),
                str(vlan.status),
                ", ".join(str_tags) if len(str_tags) else "",
            ]
        )

    # saving to file
    workbook = xlsxwriter.Workbook(report_file)
    worksheet = workbook.add_worksheet()
    worksheet.add_table(
        0,
        0,
        len(excel_data),
        len(header) - 1,
        {"columns": header, "data": excel_data},
    )
    workbook.close()
