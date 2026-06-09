# -*- coding: utf-8 -*-
#############################################################################
#    A part of Open HRMS Project <https://www.openhrms.com>
#
#    Cybrosys Technologies Pvt. Ltd.
#
#    Copyright (C) 2023-TODAY Cybrosys Technologies(<https://www.cybrosys.com>)
#    Author: Cybrosys Techno Solutions(<https://www.cybrosys.com>)
#
#    You can modify it under the terms of the GNU LESSER
#    GENERAL PUBLIC LICENSE (LGPL v3), Version 3.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU LESSER GENERAL PUBLIC LICENSE (LGPL v3) for more details.
#
#    You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENSE
#    (LGPL v3) along with this program.
#    If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################
{
    'name': "ERP Dashboard",
    'version': '17.0.1.0.2',
    'summary': """Open HRMS - HR Dashboard""",
    'description': """Open HRMS - HR Dashboard""",
    'category': 'Generic Modules/Human Resources',
    'live_test_url': 'https://youtu.be/XwGGvZbv6sc',
    'author': 'RedRuby Technologies',
    'company': 'RedRuby Technologies',
    'maintainer': 'RedRuby Technologies',
    'website': 'https://redrubytechnologies.com/',
    'depends': ['hr', 'web', 'hr_timesheet', 'hr_timesheet_attendance', 'stock', 'purchase', 'sale', 'project', 'zigma_erp'],
    'external_dependencies': {
        'python': ['pandas', 'xlsxwriter'],
    },
    'data': [
        'security/ir.model.access.csv',
        # 'report/broadfactor.xml',
        # 'views/hr_leave_views.xml',
        # 'report/print_label_multi_report.xml',
        # 'views/erp_dashboard_menus.xml',
        # 'views/hp_dashboard_menus.xml',
        # 'views/hrms_dashboard_menus.xml',
        'views/reports_wizard_views.xml',
        # 'views/erp_index_template.xml',
        
    ],
    'assets': {
        'web.assets_backend': [


            # 'hrms_dashboard/static/src/css/erp_dashboard.css',
            # 'hrms_dashboard/static/src/css/hrms_dashboard.css',
            # 'hrms_dashboard/static/src/css/lib/nv.d3.css',
            # 'hrms_dashboard/static/src/js/erp_dashboard.js',
            # 'hrms_dashboard/static/src/js/rev_dashboard.js',
            # 'hrms_dashboard/static/src/erp_index.js',
            # 'hrms_dashboard/static/src/js/hrms_dashboard.js',

            'hrms_dashboard/static/src/js/lib/d3.min.js',
            # 'hrms_dashboard/static/src/css/erp_dashboard.css',
            # 'hrms_dashboard/static/src/xml/erp_dashboard.xml',
            # 'hrms_dashboard/static/src/xml/rev_dashboard.xml',
            # 'hrms_dashboard/static/src/xml/hrms_dashboard.xml',
            
            # 'hrms_dashboard/static/src/js/hp_dashboard.js',
            # 'hrms_dashboard/static/src/xml/hp_dashboard.xml',
            # 'hrms_dashboard/static/src/css/hp_dashboard.scss',
            'hrms_dashboard/static/src/js/report_loader.js',
            'https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js',
        ],
    },
    'images': ["static/description/banner.jpg"],
    'license': 'LGPL-3',
    'installable': True,
    'application': True,
}
