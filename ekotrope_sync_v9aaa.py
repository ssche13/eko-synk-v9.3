"""
DSLD Homes - Ekotrope Sync App v9
ENERGY STAR 3.2 Compliant with REM/Rate Integration

Features: All v8 features plus REM file import/export, 8 chart types, extended calculators

pip install pandas openpyxl matplotlib requests reportlab pillow lxml
python ekotrope_sync_v9.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
from datetime import datetime, timedelta
import os, math
import xml.etree.ElementTree as ET
from xml.dom import minidom

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import matplotlib
    matplotlib.use('TkAgg')
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".dsld_ekotrope")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")

def ensure_config_dir():
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)

class DSLDSchema:
    PROJECT_FIELDS = ['Region', 'Subdivision1', 'Lot1', 'StreetAddress', 'City', 'State', 'ZipCode', 'Plan1', 'Living', 'PermitNo1']
    DATE_FIELDS = ['PDWCreated1', 'FinalCreatedDate', 'FinalizationDate', 'ConstCompleteDate', 'TargetClosingDate', 'ActualClosingDate']
    PERSONNEL_FIELDS = ['Super', 'Tech', 'RTIN']
    STATUS_FIELDS = ['PDWFails1', 'PassFail1']
    HVAC_FIELDS = ['ElecOption', 'SupplierName', 'Tonnage', 'RefrigeratorModel', 'RangeModel']
    DUCT_FIELDS = ['TDLCFM', 'LTOCFM', 'BDCFM', 'MVCFM', 'ReturnCount']
    AIRFLOW_FIELDS = ['ReturnIWC', 'SupplyIWC', 'BlowerCFM', 'MeasuredCFM', 'FWD', 'MeasuredWattage', 'Charge']
    BATH_FAN_FIELDS = ['BathFan1CFM', 'BathFan2CFM', 'BathFan3CFM', 'BathFanPass']
    ALL_FIELDS = PROJECT_FIELDS + DATE_FIELDS + PERSONNEL_FIELDS + STATUS_FIELDS + HVAC_FIELDS + DUCT_FIELDS + AIRFLOW_FIELDS + BATH_FAN_FIELDS
    
    @classmethod
    def get_template_fields(cls):
        return cls.PROJECT_FIELDS + ['PermitNo1', 'RTIN']

class REMFileHandler:
    @classmethod
    def read_rem_file(cls, filepath):
        projects = []
        ext = os.path.splitext(filepath)[1].lower()
        if ext == '.xml':
            projects = cls._parse_rem_xml(filepath)
        elif ext == '.csv' and HAS_PANDAS:
            df = pd.read_csv(filepath)
            for record in df.to_dict('records'):
                project = {k: v for k, v in record.items() if pd.notna(v)}
                if project:
                    projects.append(project)
        return projects
    
    @classmethod
    def _parse_rem_xml(cls, filepath):
        projects = []
        tree = ET.parse(filepath)
        root = tree.getroot()
        xml_map = {'ConditionedFloorArea': 'Living', 'TotalDuctLeakage': 'TDLCFM', 'DuctLeakageToOutside': 'LTOCFM',
                   'BlowerDoorCFM50': 'BDCFM', 'CoolingCapacity': 'Tonnage', 'SystemAirflow': 'MeasuredCFM',
                   'ReturnStaticPressure': 'ReturnIWC', 'SupplyStaticPressure': 'SupplyIWC', 'RefrigerantCharge': 'Charge'}
        for elem_name in ['Building', 'Home', 'Project', 'Rating']:
            for building in root.findall(f'.//{elem_name}'):
                project = {}
                for child in building:
                    tag = child.tag.split('}')[-1]
                    if tag in xml_map and child.text:
                        try:
                            project[xml_map[tag]] = float(child.text.strip())
                        except:
                            project[xml_map[tag]] = child.text.strip()
                if project:
                    projects.append(project)
        return projects
    
    @classmethod
    def export_to_rem_xml(cls, projects, filepath):
        root = ET.Element('REMRateExport')
        root.set('version', '1.0')
        root.set('exportDate', datetime.now().isoformat())
        root.set('source', 'DSLD Ekotrope Sync v9')
        
        for i, p in enumerate(projects):
            if not p: continue
            building = ET.SubElement(root, 'Building')
            building.set('id', str(i + 1))
            
            if p.get('StreetAddress'):
                addr = ET.SubElement(building, 'Address')
                for field, tag in [('StreetAddress', 'Street'), ('City', 'City'), ('State', 'State'), ('ZipCode', 'ZipCode')]:
                    if p.get(field):
                        ET.SubElement(addr, tag).text = str(p[field])
            
            info = ET.SubElement(building, 'BuildingInfo')
            for field, tag in [('Living', 'ConditionedFloorArea'), ('Subdivision1', 'Subdivision'), ('Lot1', 'LotNumber')]:
                if p.get(field):
                    ET.SubElement(info, tag).text = str(p[field])
            
            ducts = ET.SubElement(building, 'DuctTesting')
            if p.get('TDLCFM') is not None:
                ET.SubElement(ducts, 'TotalDuctLeakage').text = f"{p['TDLCFM']:.1f}"
            if p.get('LTOCFM') is not None:
                ET.SubElement(ducts, 'DuctLeakageToOutside').text = f"{p['LTOCFM']:.1f}"
            
            if p.get('BDCFM') is not None:
                infiltration = ET.SubElement(building, 'Infiltration')
                ET.SubElement(infiltration, 'BlowerDoorCFM50').text = f"{p['BDCFM']:.1f}"
            
            hvac = ET.SubElement(building, 'HVAC')
            for field, tag in [('Tonnage', 'CoolingCapacity'), ('MeasuredCFM', 'SystemAirflow'), ('ReturnIWC', 'ReturnStaticPressure'),
                              ('SupplyIWC', 'SupplyStaticPressure'), ('Charge', 'RefrigerantCharge')]:
                if p.get(field) is not None:
                    ET.SubElement(hvac, tag).text = f"{p[field]}"
        
        xml_str = ET.tostring(root, encoding='unicode')
        dom = minidom.parseString(xml_str)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(dom.toprettyxml(indent="  "))
        return True
    
    @classmethod
    def export_to_rem_csv(cls, projects, filepath):
        if not HAS_PANDAS:
            raise Exception("pandas required for CSV export")
        data = []
        for p in projects:
            if not p: continue
            ach50 = (p['BDCFM'] * 60) / (p['Living'] * 8) if p.get('BDCFM') and p.get('Living') and p['Living'] > 0 else None
            row = {'Subdivision': p.get('Subdivision1', ''), 'Lot': p.get('Lot1', ''), 'Address': p.get('StreetAddress', ''),
                   'Conditioned Floor Area': p.get('Living', ''), 'Total Duct Leakage CFM25': p.get('TDLCFM', ''),
                   'Leakage to Outside CFM25': p.get('LTOCFM', ''), 'Blower Door CFM50': p.get('BDCFM', ''),
                   'ACH50': f"{ach50:.2f}" if ach50 else '', 'Cooling Tons': p.get('Tonnage', ''),
                   'Pass/Fail': p.get('PassFail1', '')}
            data.append(row)
        pd.DataFrame(data).to_csv(filepath, index=False)
        return True

class ComplianceStandards:
    ENERGY_STAR_32_CZ2 = {'name': 'ENERGY STAR 3.2 CZ2', 'duct_leakage_total_rate': 8.0, 'duct_leakage_total_min': 80.0,
        'duct_leakage_total_rate_alt': 12.0, 'duct_leakage_total_min_alt': 120.0, 'duct_leakage_outside_rate': 4.0,
        'duct_leakage_outside_min': 40.0, 'lto_waiver_rate': 4.0, 'lto_waiver_min': 40.0, 'return_iwc_max': 0.20,
        'supply_iwc_max': 0.25, 'charge_tolerance': 0.05, 'cfm_per_ton_min': 350, 'cfm_per_ton_max': 450,
        'bath_fan_intermittent_min': 50, 'return_count_threshold': 3}
    
    @classmethod
    def get_standard(cls, name):
        return cls.ENERGY_STAR_32_CZ2
    
    @classmethod
    def get_all_versions(cls):
        return ['ENERGY STAR 3.0', 'ENERGY STAR 3.1', 'ENERGY STAR 3.2', 'ENERGY STAR 3.3']

class HomeOrientation:
    ORIENTATIONS = [('N', 0, 'North'), ('NE', 45, 'Northeast'), ('E', 90, 'East'), ('SE', 135, 'Southeast'),
                    ('S', 180, 'South'), ('SW', 225, 'Southwest'), ('W', 270, 'West'), ('NW', 315, 'Northwest')]
    @classmethod
    def get_all(cls):
        return [(o[0], o[2]) for o in cls.ORIENTATIONS]

class ComplianceChecker:
    def __init__(self, standard):
        self.standard = standard
    
    def check_project(self, project):
        if not project:
            return {'overall': 'FAIL', 'checks': [], 'pass_count': 0, 'fail_count': 0, 'warn_count': 0, 'footnotes_applied': []}
        results = {'overall': 'PASS', 'checks': [], 'pass_count': 0, 'fail_count': 0, 'warn_count': 0, 'footnotes_applied': []}
        living = project.get('Living') or 0
        return_count = project.get('ReturnCount') or 0
        use_fn41 = return_count >= 3
        if use_fn41:
            results['footnotes_applied'].append('Footnote 41: 3+ returns')
        
        # TDL Check
        tdl = project.get('TDLCFM')
        if tdl is not None and living > 0:
            rate = 12.0 if use_fn41 else 8.0
            minimum = 120.0 if use_fn41 else 80.0
            allowable = max((living / 100) * rate, minimum)
            status = 'PASS' if tdl <= allowable else 'FAIL'
            results['checks'].append({'component': 'Total Duct Leakage (6.4.2)', 'value': f"{tdl:.0f} CFM25",
                                     'requirement': f"<={allowable:.0f} CFM25", 'status': status})
        
        # LTO Check
        lto = project.get('LTOCFM')
        if lto is not None and living > 0:
            allowable = max((living / 100) * 4.0, 40.0)
            status = 'PASS' if lto <= allowable else 'FAIL'
            results['checks'].append({'component': 'Duct Leakage to Outside (6.5)', 'value': f"{lto:.0f} CFM25",
                                     'requirement': f"<={allowable:.0f} CFM25", 'status': status})
        
        # Static Pressure
        ret_iwc = project.get('ReturnIWC')
        if ret_iwc is not None:
            status = 'PASS' if abs(ret_iwc) <= 0.20 else 'WARN'
            results['checks'].append({'component': 'Return Static (5b.2)', 'value': f"{ret_iwc:.3f} IWC",
                                     'requirement': "<=0.20 IWC", 'status': status})
        
        sup_iwc = project.get('SupplyIWC')
        if sup_iwc is not None:
            status = 'PASS' if abs(sup_iwc) <= 0.25 else 'WARN'
            results['checks'].append({'component': 'Supply Static (5b.2)', 'value': f"{sup_iwc:.3f} IWC",
                                     'requirement': "<=0.25 IWC", 'status': status})
        
        # Charge
        charge = project.get('Charge')
        if charge is not None:
            status = 'PASS' if abs(charge) <= 0.05 else 'WARN'
            results['checks'].append({'component': 'Refrigerant Charge (5a.3)', 'value': f"{charge:.3f}",
                                     'requirement': "+/-0.05", 'status': status})
        
        # CFM/Ton
        cfm = project.get('MeasuredCFM')
        tons = project.get('Tonnage')
        if cfm and tons and tons > 0:
            cpt = cfm / tons
            status = 'PASS' if 350 <= cpt <= 450 else 'WARN' if 300 <= cpt <= 500 else 'FAIL'
            results['checks'].append({'component': 'Airflow (5a.1)', 'value': f"{cpt:.0f} CFM/ton",
                                     'requirement': "350-450 CFM/ton", 'status': status})
        
        # Bath Fan
        mvcfm = project.get('MVCFM')
        if mvcfm is not None:
            status = 'PASS' if mvcfm >= 50 else 'FAIL'
            results['checks'].append({'component': 'Bath Fan (8.2)', 'value': f"{mvcfm:.0f} CFM",
                                     'requirement': ">=50 CFM", 'status': status})
        
        # Tally
        for c in results['checks']:
            if c['status'] == 'PASS': results['pass_count'] += 1
            elif c['status'] == 'FAIL':
                results['fail_count'] += 1
                results['overall'] = 'FAIL'
            elif c['status'] == 'WARN':
                results['warn_count'] += 1
                if results['overall'] == 'PASS': results['overall'] = 'WARN'
        return results

class DataValidator:
    def validate_project(self, project):
        issues = {'errors': [], 'warnings': [], 'info': [], 'is_valid': False, 'total_issues': 0}
        if not project:
            issues['errors'].append("No project data")
            return issues
        if not project.get('Subdivision1'): issues['errors'].append("Missing: Subdivision")
        if not project.get('Lot1'): issues['errors'].append("Missing: Lot")
        if not project.get('Living') or project.get('Living', 0) < 500: issues['errors'].append("Missing/invalid: Living sqft")
        if not project.get('StreetAddress'): issues['errors'].append("Missing: Address")
        if project.get('TDLCFM') is None: issues['warnings'].append("Missing: TDLCFM")
        if project.get('LTOCFM') is None: issues['warnings'].append("Missing: LTOCFM")
        pf = project.get('PassFail1')
        if pf and str(pf).lower() == 'fail': issues['errors'].append("Project marked FAIL")
        issues['is_valid'] = len(issues['errors']) == 0
        issues['total_issues'] = len(issues['errors']) + len(issues['warnings'])
        return issues

class RatingType:
    @classmethod
    def determine(cls, project):
        has_final = project.get('FinalCreatedDate') is not None
        passed = str(project.get('PassFail1', '')).lower() == 'pass'
        return 'Confirmed' if has_final and passed else 'Projected'

class EkotropeJSONGenerator:
    def __init__(self, config):
        self.config = config
    
    def generate(self, projects, target_version='ENERGY STAR 3.2', orientation='N'):
        homes = []
        for p in projects:
            if not p: continue
            template = self.config.get('builder_home_id_template', '{Subdivision1}_Lot{Lot1}')
            builder_id = template
            for field in DSLDSchema.get_template_fields():
                builder_id = builder_id.replace('{' + field + '}', str(p.get(field, '') or '').strip().replace(' ', '_'))
            if not builder_id or builder_id == '_Lot': continue
            home = {'builderHomeId': builder_id, 'ratingType': RatingType.determine(p), 'targetEnergyStarVersion': target_version}
            if p.get('StreetAddress'):
                home['address'] = {'street': str(p.get('StreetAddress', '')), 'city': str(p.get('City', '')),
                                  'state': str(p.get('State', '')), 'zip': str(p.get('ZipCode', ''))}
            home['generalInfo'] = {'conditionedFloorArea': p.get('Living'), 'orientation': orientation}
            if p.get('BDCFM') is not None:
                home['infiltration'] = {'value': float(p['BDCFM']), 'unit': 'CFM50'}
            if p.get('TDLCFM') is not None or p.get('LTOCFM') is not None:
                dist = {'index': 0}
                if p.get('TDLCFM') is not None: dist['totalDuctLeakageCfm25'] = float(p['TDLCFM'])
                if p.get('LTOCFM') is not None: dist['leakageToOutsideCfm25'] = float(p['LTOCFM'])
                home['distributionSystems'] = [dist]
            homes.append(home)
        return {'homes': homes, 'metadata': {'generated': datetime.now().isoformat(), 'source': 'DSLD v9', 'count': len(homes)}}

class ConstructionCalculators:
    @staticmethod
    def duct_leakage_per_100(cfm, sqft): return (cfm or 0) / sqft * 100 if sqft and sqft > 0 else 0
    @staticmethod
    def allowable_duct_leakage(sqft, rate=8.0, minimum=80.0): return max((sqft / 100) * rate, minimum) if sqft and sqft > 0 else minimum
    @staticmethod
    def cfm_per_ton(cfm, tons): return (cfm or 0) / tons if tons and tons > 0 else 0
    @staticmethod
    def total_external_sp(ret_iwc, sup_iwc): return abs(ret_iwc or 0) + abs(sup_iwc or 0)
    @staticmethod
    def ach50(bdcfm, sqft, ceiling_height=8): return (bdcfm or 0) * 60 / (sqft * ceiling_height) if sqft and sqft > 0 else 0
    @staticmethod
    def recommended_tonnage(sqft, factor=500): return sqft / factor if sqft and sqft > 0 else 0
    @staticmethod
    def natural_ach(ach50, n=17): return ach50 / n if ach50 else 0
    @staticmethod
    def required_ventilation_cfm(sqft, br): return 0.01 * (sqft or 0) + 7.5 * ((br or 2) + 1)

class ThemeManager:
    LIGHT = {'name': 'Light', 'bg': '#f5f5f5', 'fg': '#1a1a1a', 'bg_alt': '#ffffff', 'accent': '#1e5799',
             'success': '#1e7e34', 'warning': '#d39e00', 'error': '#c82333', 'border': '#cccccc',
             'tree_bg': '#ffffff', 'tree_fg': '#1a1a1a', 'tree_selected': '#1e5799', 'input_bg': '#ffffff',
             'button_bg': '#e0e0e0', 'header_bg': '#1e5799', 'header_fg': '#ffffff',
             'pass_bg': '#c3e6cb', 'fail_bg': '#f5c6cb', 'warn_bg': '#ffeeba', 'info_bg': '#d1ecf1'}
    DARK = {'name': 'Dark', 'bg': '#1a1a2e', 'fg': '#eaeaea', 'bg_alt': '#16213e', 'accent': '#4da6ff',
            'success': '#28a745', 'warning': '#ffc107', 'error': '#dc3545', 'border': '#3a3a5c',
            'tree_bg': '#16213e', 'tree_fg': '#eaeaea', 'tree_selected': '#4da6ff', 'input_bg': '#0f3460',
            'button_bg': '#0f3460', 'header_bg': '#0f3460', 'header_fg': '#ffffff',
            'pass_bg': '#155724', 'fail_bg': '#721c24', 'warn_bg': '#856404', 'info_bg': '#0c5460'}
    def __init__(self): self.current = self.DARK
    def toggle(self): self.current = self.LIGHT if self.current == self.DARK else self.DARK; return self.current
    def get(self, key): return self.current.get(key, '#000000')
    def is_dark(self): return self.current == self.DARK

class ConfigManager:
    DEFAULT = {'builder_home_id_template': '{Subdivision1}_Lot{Lot1}', 'target_energy_star_version': 'ENERGY STAR 3.2',
               'current_user': '', 'theme': 'dark', 'default_orientation': 'N'}
    def __init__(self):
        ensure_config_dir()
        self.config = self._load()
    def _load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f: return {**self.DEFAULT, **json.load(f)}
            except: pass
        return self.DEFAULT.copy()
    def save(self):
        ensure_config_dir()
        with open(CONFIG_FILE, 'w') as f: json.dump(self.config, f, indent=2)
    def get(self, key, default=None): return self.config.get(key, default)
    def set(self, key, value): self.config[key] = value; self.save()

class ExcelLoader:
    # Column alias map: maps various SQL export names -> internal standard name
    # Supports multiple export formats (with/without numbers, spaces, etc.)
    COLUMN_ALIASES = {
        'Region':           ['Region'],
        'Subdivision1':     ['Subdivision1', 'Subdivision', 'Sub', 'SubdivisionName'],
        'Lot1':             ['Lot1', 'Lot', 'LotNumber', 'Lot Number', 'LotNo'],
        'StreetAddress':    ['StreetAddress', 'Street Address', 'Address', 'Street'],
        'City':             ['City'],
        'State':            ['State'],
        'ZipCode':          ['ZipCode', 'Zip Code', 'Zip', 'PostalCode'],
        'Plan1':            ['Plan1', 'Plan', 'PlanName'],
        'Living':           ['Living', 'LivingSqFt', 'Living SqFt'],
        'PermitNo1':        ['PermitNo1', 'Permit No', 'PermitNo', 'Permit Number'],
        'ConstCompleteDate1': ['ConstCompleteDate1', 'Const Complete Date', 'ConstCompleteDate'],
        'Super':            ['Super', 'Superintendent'],
        'Tech':             ['Tech', 'Technician'],
        'RTIN':             ['RTIN'],
        'PDWCreated1':      ['PDWCreated1', 'PDWCreated', 'PDW Created'],
        'PDWFails1':        ['PDWFails1', 'PDWFails', 'PDW Fails'],
        'FinalCreatedDate': ['FinalCreatedDate', 'Final Created Date', 'FinalCreated'],
        'FinalizationDate': ['FinalizationDate', 'Finalization Date'],
        'PassFail1':        ['PassFail1', 'PassFail', 'Pass Fail', 'Final Fails'],
        'ConstCompleteDate': ['ConstCompleteDate', 'Const Complete Date'],
        'TargetClosingDate': ['TargetClosingDate', 'Target Closing Date', 'TargetClosing'],
        'ActualClosingDate': ['ActualClosingDate', 'Actual Closing Date', 'ActualClosing'],
        'ElecOption':       ['ElecOption', 'Elec Option', 'Electric Option'],
        'SupplierName':     ['SupplierName', 'Supplier Name', 'HVAC', 'HVACSupplier'],
        'Tonnage':          ['Tonnage', 'Tons'],
        'RefrigeratorModel': ['RefrigeratorModel', 'Refrigerator Model'],
        'RangeModel':       ['RangeModel', 'Range Model'],
        'TDLCFM':           ['TDLCFM', 'TDL CFM', 'TotalDuctLeakage'],
        'LTOCFM':           ['LTOCFM', 'LTO CFM', 'LeakageToOutside'],
        'BDCFM':            ['BDCFM', 'BD CFM', 'BlowerDoor'],
        'MVCFM':            ['MVCFM', 'MV CFM', 'MasterVent'],
        'ReturnIWC':        ['ReturnIWC', 'Return IWC', 'ReturnStaticPressure'],
        'SupplyIWC':        ['SupplyIWC', 'Supply IWC', 'SupplyStaticPressure'],
        'BlowerCFM':        ['BlowerCFM', 'Blower CFM', 'BlowerAirflow'],
        'MeasuredCFM':      ['MeasuredCFM', 'Measured CFM', 'MeasuredAirflow'],
        'FWD':              ['FWD', 'FanWattDraw', 'Fan Watt Draw'],
        'MeasuredWattage':  ['MeasuredWattage', 'Measured Wattage'],
        'Charge':           ['Charge', 'RefrigerantCharge', 'Refrigerant Charge'],
    }
    
    @staticmethod
    def _normalize_columns(df):
        """Normalize column names to internal standard names.
        Handles multiple SQL export formats (with/without numbers, spaces, etc.)"""
        # Build reverse lookup: any alias (lowercase, no spaces) -> standard name
        alias_map = {}
        for standard, aliases in ExcelLoader.COLUMN_ALIASES.items():
            for alias in aliases:
                alias_map[alias.lower().replace(' ', '').strip()] = standard
        
        rename_map = {}
        for col in df.columns:
            clean = str(col).strip()
            lookup = clean.lower().replace(' ', '')
            if lookup in alias_map:
                target = alias_map[lookup]
                if clean != target:
                    rename_map[col] = target
            elif clean != col:
                rename_map[col] = clean  # At least strip whitespace
        
        if rename_map:
            df = df.rename(columns=rename_map)
        return df
    
    @staticmethod
    def load_file(filepath):
        if not HAS_PANDAS: raise Exception("pandas not installed")
        df = pd.read_excel(filepath)
        
        # Normalize column names (handles extra spaces, different casing)
        df = ExcelLoader._normalize_columns(df)
        
        projects = df.to_dict('records')
        for p in projects:
            for k, v in list(p.items()):
                if pd.isna(v): p[k] = None
                elif isinstance(v, pd.Timestamp): p[k] = v.strftime('%Y-%m-%d')
        return projects

class EkotropeSyncApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DSLD Homes - Ekotrope Sync v9 (REM/Rate Integration)")
        self.root.geometry("1500x950")
        self.config = ConfigManager()
        self.theme = ThemeManager()
        if self.config.get('theme') == 'light': self.theme.current = ThemeManager.LIGHT
        self.json_gen = EkotropeJSONGenerator(self.config)
        self.calc = ConstructionCalculators()
        self.validator = DataValidator()
        self.all_projects = {}
        self.validation_results = {}
        self.compliance_results = {}
        self.current_user = self.config.get('current_user', 'Unknown')
        self._apply_theme()
        self._build_ui()
        if not self.current_user or self.current_user == 'Unknown': self._prompt_user()
    
    def _apply_theme(self):
        style = ttk.Style()
        t = self.theme.current
        style.configure('.', background=t['bg'], foreground=t['fg'])
        style.configure('TFrame', background=t['bg'])
        style.configure('TLabel', background=t['bg'], foreground=t['fg'])
        style.configure('TLabelframe', background=t['bg'], foreground=t['fg'])
        style.configure('TLabelframe.Label', background=t['bg'], foreground=t['fg'], font=('Arial', 10, 'bold'))
        style.configure('TButton', background=t['button_bg'])
        style.configure('TEntry', fieldbackground=t['input_bg'])
        style.configure('TCombobox', fieldbackground=t['input_bg'])
        style.configure('Treeview', background=t['tree_bg'], foreground=t['tree_fg'], fieldbackground=t['tree_bg'])
        style.map('Treeview', background=[('selected', t['tree_selected'])])
        style.configure('TNotebook', background=t['bg'])
        style.configure('TNotebook.Tab', background=t['button_bg'], padding=[12, 6])
        style.map('TNotebook.Tab', background=[('selected', t['accent'])])
        style.configure('Header.TFrame', background=t['header_bg'])
        style.configure('Header.TLabel', background=t['header_bg'], foreground=t['header_fg'], font=('Arial', 16, 'bold'))
        style.configure('HeaderSub.TLabel', background=t['header_bg'], foreground=t['header_fg'], font=('Arial', 10))
        self.root.configure(bg=t['bg'])
    
    def _toggle_theme(self):
        self.theme.toggle()
        self.config.set('theme', 'light' if not self.theme.is_dark() else 'dark')
        self._apply_theme()
        self.theme_btn.config(text=" Light" if self.theme.is_dark() else " Dark")
        self._update_tree_tags()
        if HAS_MATPLOTLIB: self.refresh_charts()
    
    def _update_tree_tags(self):
        t = self.theme.current
        self.tree.tag_configure('pass', background=t['pass_bg'])
        self.tree.tag_configure('fail', background=t['fail_bg'])
        self.tree.tag_configure('warn', background=t['warn_bg'])
        self.tree.tag_configure('info', background=t['info_bg'])
    
    def _prompt_user(self):
        name = simpledialog.askstring("User", "Enter your name:", parent=self.root)
        if name:
            self.config.set('current_user', name.strip())
            self.current_user = name.strip()
            self.user_lbl.config(text=f" {self.current_user}")
    
    def _build_ui(self):
        # Header
        header = ttk.Frame(self.root, style='Header.TFrame')
        header.pack(fill='x')
        header_inner = ttk.Frame(header, style='Header.TFrame')
        header_inner.pack(fill='x', padx=20, pady=8)
        title_frame = ttk.Frame(header_inner, style='Header.TFrame')
        title_frame.pack(side='left')
        ttk.Label(title_frame, text="DSLD HOMES", style='Header.TLabel').pack(anchor='w')
        ttk.Label(title_frame, text="Ekotrope Sync v9 - REM/Rate Integration", style='HeaderSub.TLabel').pack(anchor='w')
        right_frame = ttk.Frame(header_inner, style='Header.TFrame')
        right_frame.pack(side='right')
        self.user_lbl = ttk.Label(right_frame, text=f" {self.current_user}", style='HeaderSub.TLabel')
        self.user_lbl.pack(side='right', padx=10)
        self.theme_btn = ttk.Button(right_frame, text=" Light" if self.theme.is_dark() else " Dark", command=self._toggle_theme, width=10)
        self.theme_btn.pack(side='right', padx=5)
        
        # Menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        fm = tk.Menu(menubar, tearoff=0)
        fm.add_command(label="Load Excel File...", command=self.load_excel_file)
        fm.add_command(label="Load REM/Rate File...", command=self.load_rem_file)
        fm.add_separator()
        fm.add_command(label="Export to JSON...", command=self.generate_json)
        fm.add_command(label="Export to REM XML...", command=self.export_rem_xml)
        fm.add_command(label="Export to REM CSV...", command=self.export_rem_csv)
        fm.add_separator()
        fm.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=fm)
        sm = tk.Menu(menubar, tearoff=0)
        sm.add_command(label="Configure Template...", command=self.configure_template)
        sm.add_command(label="Change User...", command=self._prompt_user)
        menubar.add_cascade(label="Settings", menu=sm)
        hm = tk.Menu(menubar, tearoff=0)
        hm.add_command(label="About", command=self.show_about)
        hm.add_command(label="ENERGY STAR 3.2 Reference", command=self.show_compliance_ref)
        menubar.add_cascade(label="Help", menu=hm)
        
        # Notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        self.export_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.export_tab, text="   Export  ")
        self._build_export_tab()
        self.validation_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.validation_tab, text="   Validation  ")
        self._build_validation_tab()
        self.compliance_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.compliance_tab, text="   Compliance  ")
        self._build_compliance_tab()
        self.charts_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.charts_tab, text="   Charts  ")
        self._build_charts_tab()
        self.calc_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.calc_tab, text="   Calculators  ")
        self._build_calc_tab()
        self.rem_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.rem_tab, text="   REM/Rate  ")
        self._build_rem_tab()
        
        # Status bar
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill='x', side='bottom')
        self.status = ttk.Label(status_frame, text="Ready - Load an Excel or REM file to begin", relief='sunken')
        self.status.pack(fill='x', side='left', expand=True)
        self.count_lbl = ttk.Label(status_frame, text="0 projects", relief='sunken', width=15)
        self.count_lbl.pack(side='right')
    
    def _build_export_tab(self):
        main = ttk.Frame(self.export_tab)
        main.pack(fill='both', expand=True, padx=10, pady=10)
        top = ttk.Frame(main)
        top.pack(fill='x', pady=(0, 10))
        src_frame = ttk.LabelFrame(top, text="Data Source")
        src_frame.pack(side='left', fill='x', expand=True, padx=(0, 10))
        self.source_lbl = ttk.Label(src_frame, text="No file loaded", font=('Arial', 10, 'bold'))
        self.source_lbl.pack(side='left', padx=10, pady=5)
        ttk.Button(src_frame, text=" Load Excel...", command=self.load_excel_file).pack(side='right', padx=5, pady=5)
        ttk.Button(src_frame, text=" Load REM...", command=self.load_rem_file).pack(side='right', padx=5, pady=5)
        settings_frame = ttk.LabelFrame(top, text="Export Settings")
        settings_frame.pack(side='left', padx=(0, 10))
        ttk.Label(settings_frame, text="Version:").grid(row=0, column=0, padx=5, pady=2, sticky='e')
        self.version_cb = ttk.Combobox(settings_frame, values=ComplianceStandards.get_all_versions(), width=18, state='readonly')
        self.version_cb.set(self.config.get('target_energy_star_version', 'ENERGY STAR 3.2'))
        self.version_cb.grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(settings_frame, text="Orientation:").grid(row=1, column=0, padx=5, pady=2, sticky='e')
        self.orientation_cb = ttk.Combobox(settings_frame, values=[f"{o[0]} - {o[1]}" for o in HomeOrientation.get_all()], width=18, state='readonly')
        self.orientation_cb.set(f"{self.config.get('default_orientation', 'N')} - North")
        self.orientation_cb.grid(row=1, column=1, padx=5, pady=2)
        
        filter_frame = ttk.LabelFrame(main, text="Filters")
        filter_frame.pack(fill='x', pady=(0, 10))
        filter_row = ttk.Frame(filter_frame)
        filter_row.pack(fill='x', padx=5, pady=5)
        ttk.Label(filter_row, text="Region:").pack(side='left', padx=5)
        self.region_cb = ttk.Combobox(filter_row, width=15, state='readonly')
        self.region_cb.pack(side='left', padx=5)
        self.region_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())
        ttk.Label(filter_row, text="Status:").pack(side='left', padx=5)
        self.status_cb = ttk.Combobox(filter_row, values=['All', 'Pass', 'Fail'], width=8, state='readonly')
        self.status_cb.set('All')
        self.status_cb.pack(side='left', padx=5)
        self.status_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())
        ttk.Button(filter_row, text="Clear Filters", command=self.clear_filters).pack(side='right', padx=5)
        
        list_frame = ttk.LabelFrame(main, text="Projects")
        list_frame.pack(fill='both', expand=True, pady=(0, 10))
        cols = ('lot', 'address', 'subdivision', 'sqft', 'tons', 'tdl', 'lto', 'bd', 'rating', 'pf')
        self.tree = ttk.Treeview(list_frame, columns=cols, show='headings', selectmode='extended')
        for c, w in [('lot', 50), ('address', 150), ('subdivision', 120), ('sqft', 55), ('tons', 45), ('tdl', 50), ('lto', 50), ('bd', 60), ('rating', 70), ('pf', 35)]:
            self.tree.heading(c, text=c.title())
            self.tree.column(c, width=w)
        self._update_tree_tags()
        vsb = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill='x')
        self.sel_lbl = ttk.Label(btn_frame, text="Selected: 0 of 0")
        self.sel_lbl.pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Select All", command=self.select_all).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="Preview JSON", command=self.preview_json).pack(side='left', padx=2)
        ttk.Button(btn_frame, text=" Export JSON", command=self.generate_json).pack(side='left', padx=5)
    
    def _build_validation_tab(self):
        top = ttk.Frame(self.validation_tab)
        top.pack(fill='x', padx=10, pady=10)
        ttk.Button(top, text=" Validate All", command=self.run_validation).pack(side='left', padx=5)
        self.val_sum = ttk.Label(top, text="Click 'Validate All' to check", font=('Arial', 10))
        self.val_sum.pack(side='left', padx=20)
        paned = ttk.PanedWindow(self.validation_tab, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=10, pady=5)
        left = ttk.Frame(paned)
        paned.add(left, weight=1)
        self.val_tree = ttk.Treeview(left, columns=('project', 'errors', 'warnings', 'status'), show='headings')
        for c in ['project', 'errors', 'warnings', 'status']:
            self.val_tree.heading(c, text=c.title())
        self.val_tree.pack(fill='both', expand=True)
        self.val_tree.bind('<<TreeviewSelect>>', self.show_val_details)
        right = ttk.Frame(paned)
        paned.add(right, weight=2)
        self.val_txt = tk.Text(right, wrap='word', font=('Consolas', 10), bg=self.theme.get('bg_alt'), fg=self.theme.get('fg'))
        self.val_txt.pack(fill='both', expand=True)
    
    def _build_compliance_tab(self):
        top = ttk.Frame(self.compliance_tab)
        top.pack(fill='x', padx=10, pady=10)
        ttk.Label(top, text="Standard:").pack(side='left', padx=5)
        self.std_cb = ttk.Combobox(top, values=ComplianceStandards.get_all_versions(), width=18, state='readonly')
        self.std_cb.set('ENERGY STAR 3.2')
        self.std_cb.pack(side='left', padx=5)
        ttk.Button(top, text=" Check Compliance", command=self.run_compliance).pack(side='left', padx=10)
        self.comp_sum = ttk.Label(top, text="Run compliance check", font=('Arial', 10))
        self.comp_sum.pack(side='left', padx=20)
        paned = ttk.PanedWindow(self.compliance_tab, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=10, pady=5)
        left = ttk.Frame(paned)
        paned.add(left, weight=1)
        self.comp_tree = ttk.Treeview(left, columns=('project', 'pass', 'fail', 'warn', 'result'), show='headings')
        for c in ['project', 'pass', 'fail', 'warn', 'result']:
            self.comp_tree.heading(c, text=c.title())
        self.comp_tree.pack(fill='both', expand=True)
        self.comp_tree.bind('<<TreeviewSelect>>', self.show_comp_details)
        right = ttk.Frame(paned)
        paned.add(right, weight=2)
        self.comp_txt = tk.Text(right, wrap='word', font=('Consolas', 10), bg=self.theme.get('bg_alt'), fg=self.theme.get('fg'))
        self.comp_txt.pack(fill='both', expand=True)
    
    def _build_charts_tab(self):
        if not HAS_MATPLOTLIB:
            ttk.Label(self.charts_tab, text=" Charts require matplotlib\npip install matplotlib", font=('Arial', 14)).pack(pady=50)
            return
        top = ttk.Frame(self.charts_tab)
        top.pack(fill='x', padx=10, pady=5)
        ttk.Label(top, text="Chart:").pack(side='left', padx=5)
        self.chart_type_cb = ttk.Combobox(top, values=['Overview', 'Duct Leakage', 'HVAC', 'Static Pressure', 'By Region'], width=20, state='readonly')
        self.chart_type_cb.set('Overview')
        self.chart_type_cb.pack(side='left', padx=5)
        self.chart_type_cb.bind('<<ComboboxSelected>>', lambda e: self.refresh_charts())
        ttk.Button(top, text=" Refresh", command=self.refresh_charts).pack(side='left', padx=10)
        self.chart_frame = ttk.Frame(self.charts_tab)
        self.chart_frame.pack(fill='both', expand=True, padx=5, pady=5)
    
    def _build_calc_tab(self):
        nb = ttk.Notebook(self.calc_tab)
        nb.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Duct Leakage
        dl_frame = ttk.Frame(nb)
        nb.add(dl_frame, text="Duct Leakage")
        calc1 = ttk.LabelFrame(dl_frame, text="Allowable Duct Leakage")
        calc1.pack(padx=20, pady=20, anchor='nw')
        ttk.Label(calc1, text="Living Sqft:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.dl_sqft = ttk.Entry(calc1, width=12); self.dl_sqft.grid(row=0, column=1, padx=5); self.dl_sqft.insert(0, "1800")
        ttk.Label(calc1, text="Returns:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.dl_returns = ttk.Entry(calc1, width=12); self.dl_returns.grid(row=1, column=1, padx=5); self.dl_returns.insert(0, "4")
        ttk.Label(calc1, text="Measured TDL:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.dl_measured = ttk.Entry(calc1, width=12); self.dl_measured.grid(row=2, column=1, padx=5); self.dl_measured.insert(0, "80")
        ttk.Button(calc1, text="Calculate", command=self.calc_duct_limits).grid(row=3, column=0, columnspan=2, pady=10)
        self.dl_res = ttk.Label(calc1, text="", font=('Arial', 10), wraplength=300); self.dl_res.grid(row=4, column=0, columnspan=2, pady=5)
        
        # CFM/Ton
        ct_frame = ttk.Frame(nb)
        nb.add(ct_frame, text="CFM/Ton")
        calc2 = ttk.LabelFrame(ct_frame, text="Airflow per Ton")
        calc2.pack(padx=20, pady=20, anchor='nw')
        ttk.Label(calc2, text="CFM:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.ct_cfm = ttk.Entry(calc2, width=12); self.ct_cfm.grid(row=0, column=1, padx=5); self.ct_cfm.insert(0, "1400")
        ttk.Label(calc2, text="Tonnage:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.ct_ton = ttk.Entry(calc2, width=12); self.ct_ton.grid(row=1, column=1, padx=5); self.ct_ton.insert(0, "3.5")
        ttk.Button(calc2, text="Calculate", command=self.calc_cfm_per_ton).grid(row=2, column=0, columnspan=2, pady=10)
        self.ct_res = ttk.Label(calc2, text="", font=('Arial', 10)); self.ct_res.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Static Pressure
        sp_frame = ttk.Frame(nb)
        nb.add(sp_frame, text="Static Pressure")
        calc3 = ttk.LabelFrame(sp_frame, text="Total ESP")
        calc3.pack(padx=20, pady=20, anchor='nw')
        ttk.Label(calc3, text="Return IWC:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.sp_ret = ttk.Entry(calc3, width=12); self.sp_ret.grid(row=0, column=1, padx=5); self.sp_ret.insert(0, "0.15")
        ttk.Label(calc3, text="Supply IWC:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.sp_sup = ttk.Entry(calc3, width=12); self.sp_sup.grid(row=1, column=1, padx=5); self.sp_sup.insert(0, "0.20")
        ttk.Button(calc3, text="Calculate", command=self.calc_static_pressure).grid(row=2, column=0, columnspan=2, pady=10)
        self.sp_res = ttk.Label(calc3, text="", font=('Arial', 10), wraplength=300); self.sp_res.grid(row=3, column=0, columnspan=2, pady=5)
        
        # ACH50
        ach_frame = ttk.Frame(nb)
        nb.add(ach_frame, text="ACH50")
        calc4 = ttk.LabelFrame(ach_frame, text="Air Changes @ 50 Pa")
        calc4.pack(padx=20, pady=20, anchor='nw')
        ttk.Label(calc4, text="BD CFM50:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.ach_bd = ttk.Entry(calc4, width=12); self.ach_bd.grid(row=0, column=1, padx=5); self.ach_bd.insert(0, "1000")
        ttk.Label(calc4, text="Living Sqft:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.ach_sqft = ttk.Entry(calc4, width=12); self.ach_sqft.grid(row=1, column=1, padx=5); self.ach_sqft.insert(0, "1800")
        ttk.Label(calc4, text="Ceiling Ht:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.ach_height = ttk.Entry(calc4, width=12); self.ach_height.grid(row=2, column=1, padx=5); self.ach_height.insert(0, "8")
        ttk.Button(calc4, text="Calculate", command=self.calc_ach50).grid(row=3, column=0, columnspan=2, pady=10)
        self.ach_res = ttk.Label(calc4, text="", font=('Arial', 10), wraplength=300); self.ach_res.grid(row=4, column=0, columnspan=2, pady=5)
        
        # Ventilation
        vent_frame = ttk.Frame(nb)
        nb.add(vent_frame, text="Ventilation")
        calc5 = ttk.LabelFrame(vent_frame, text="ASHRAE 62.2 Requirement")
        calc5.pack(padx=20, pady=20, anchor='nw')
        ttk.Label(calc5, text="Living Sqft:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.vent_sqft = ttk.Entry(calc5, width=12); self.vent_sqft.grid(row=0, column=1, padx=5); self.vent_sqft.insert(0, "1800")
        ttk.Label(calc5, text="Bedrooms:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.vent_br = ttk.Entry(calc5, width=12); self.vent_br.grid(row=1, column=1, padx=5); self.vent_br.insert(0, "3")
        ttk.Button(calc5, text="Calculate", command=self.calc_ventilation).grid(row=2, column=0, columnspan=2, pady=10)
        self.vent_res = ttk.Label(calc5, text="", font=('Arial', 10), wraplength=300); self.vent_res.grid(row=3, column=0, columnspan=2, pady=5)
    
    def _build_rem_tab(self):
        main = ttk.Frame(self.rem_tab)
        main.pack(fill='both', expand=True, padx=10, pady=10)
        import_frame = ttk.LabelFrame(main, text="Import from REM/Rate")
        import_frame.pack(fill='x', pady=(0, 10))
        import_row = ttk.Frame(import_frame)
        import_row.pack(fill='x', padx=10, pady=10)
        ttk.Label(import_row, text="Supported: XML, CSV").pack(side='left', padx=5)
        ttk.Button(import_row, text=" Load REM File...", command=self.load_rem_file).pack(side='right', padx=5)
        export_frame = ttk.LabelFrame(main, text="Export to REM/Rate")
        export_frame.pack(fill='x', pady=(0, 10))
        export_row = ttk.Frame(export_frame)
        export_row.pack(fill='x', padx=10, pady=10)
        ttk.Button(export_row, text=" Export XML...", command=self.export_rem_xml).pack(side='left', padx=5)
        ttk.Button(export_row, text=" Export CSV...", command=self.export_rem_csv).pack(side='left', padx=5)
        map_frame = ttk.LabelFrame(main, text="Field Mapping")
        map_frame.pack(fill='both', expand=True)
        map_tree = ttk.Treeview(map_frame, columns=('dsld', 'rem', 'desc'), show='headings')
        for c, w in [('dsld', 120), ('rem', 180), ('desc', 250)]:
            map_tree.heading(c, text=c.upper())
            map_tree.column(c, width=w)
        mappings = [('Living', 'ConditionedFloorArea', 'Conditioned sqft'), ('TDLCFM', 'TotalDuctLeakage', 'Duct leak CFM25'),
                   ('LTOCFM', 'DuctLeakageToOutside', 'LTO CFM25'), ('BDCFM', 'BlowerDoorCFM50', 'Blower door'),
                   ('Tonnage', 'CoolingCapacity', 'AC tons'), ('MeasuredCFM', 'SystemAirflow', 'Airflow'),
                   ('ReturnIWC', 'ReturnStaticPressure', 'Return IWC'), ('SupplyIWC', 'SupplyStaticPressure', 'Supply IWC'),
                   ('Charge', 'RefrigerantCharge', 'Charge variance')]
        for m in mappings:
            map_tree.insert('', 'end', values=m)
        map_tree.pack(fill='both', expand=True, padx=5, pady=5)
    
    # ================================================================
    # DATA LOADING
    # ================================================================
    
    def load_excel_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath: return
        try:
            projects = ExcelLoader.load_file(filepath)
            # Generate unique keys - prevent collisions when columns are missing
            self.all_projects = {}
            missing_cols = []
            for i, p in enumerate(projects):
                if not p: continue
                sub = p.get('Subdivision1')
                lot = p.get('Lot1')
                if sub and lot:
                    key = f"{sub}_Lot{lot}"
                    if key in self.all_projects:
                        key = f"{key}_{i}"
                else:
                    addr = p.get('StreetAddress', '') or ''
                    key = f"Row{i+1}_{addr[:20]}" if addr else f"Row{i+1}"
                    if not missing_cols:
                        missing_cols = [c for c in ['Subdivision1', 'Lot1'] if not p.get(c)]
                self.all_projects[key] = p
            self.source_lbl.config(text=f" {os.path.basename(filepath)}")
            self.status.config(text=f"Loaded {len(self.all_projects)} projects from {filepath}")
            self.count_lbl.config(text=f"{len(self.all_projects)} projects")
            self._populate_filters()
            self._populate_tree()
            
            # Warn user if key columns were missing (helps troubleshoot)
            if missing_cols:
                found_cols = sorted(set(projects[0].keys())) if projects else []
                messagebox.showwarning("Column Warning",
                    f"Missing expected columns: {', '.join(missing_cols)}\n\n"
                    f"Found {len(found_cols)} columns in file.\n"
                    f"Loaded {len(self.all_projects)} projects using fallback keys.\n\n"
                    f"Check that your Excel has correct column headers.")
        except Exception as e:
            messagebox.showerror("Load Error", f"{str(e)}\n\nFile: {filepath}")
    
    def load_rem_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("REM files", "*.xml *.csv"), ("All files", "*.*")])
        if not filepath: return
        try:
            projects = REMFileHandler.read_rem_file(filepath)
            for i, p in enumerate(projects):
                key = p.get('Subdivision1', '') or f"REM_{i+1}"
                lot = p.get('Lot1', '') or str(i+1)
                self.all_projects[f"{key}_Lot{lot}"] = p
            self.source_lbl.config(text=f" {os.path.basename(filepath)}")
            self.status.config(text=f"Loaded {len(projects)} from REM file")
            self.count_lbl.config(text=f"{len(self.all_projects)} projects")
            self._populate_filters()
            self._populate_tree()
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
    
    def _populate_filters(self):
        regions = sorted(set(str(p.get('Region', 'Unknown')) for p in self.all_projects.values() if p))
        self.region_cb['values'] = ['All'] + regions
        self.region_cb.set('All')
    
    def _populate_tree(self, projects=None):
        self.tree.delete(*self.tree.get_children())
        if projects is None:
            projects = self.all_projects
        for key, p in projects.items():
            if not p: continue
            lot = str(p.get('Lot1', ''))[:8]
            addr = str(p.get('StreetAddress', ''))[:25]
            subdiv = str(p.get('Subdivision1', ''))[:18]
            sqft = f"{p.get('Living', 0):.0f}" if p.get('Living') else ''
            tons = f"{p.get('Tonnage', 0):.1f}" if p.get('Tonnage') else ''
            tdl = f"{p.get('TDLCFM', 0):.0f}" if p.get('TDLCFM') is not None else ''
            lto = f"{p.get('LTOCFM', 0):.0f}" if p.get('LTOCFM') is not None else ''
            bd = f"{p.get('BDCFM', 0):.0f}" if p.get('BDCFM') is not None else ''
            rating = RatingType.determine(p)
            pf = str(p.get('PassFail1', ''))[:4]
            tag = 'pass' if pf.lower() == 'pass' else 'fail' if pf.lower() == 'fail' else ''
            self.tree.insert('', 'end', iid=key, values=(lot, addr, subdiv, sqft, tons, tdl, lto, bd, rating, pf), tags=(tag,))
        self.sel_lbl.config(text=f"Selected: 0 of {len(projects)}")
    
    def apply_filters(self):
        region = self.region_cb.get()
        status = self.status_cb.get()
        filtered = {}
        for key, p in self.all_projects.items():
            if not p: continue
            if region != 'All' and str(p.get('Region', '')) != region: continue
            pf = str(p.get('PassFail1', '')).lower()
            if status == 'Pass' and pf != 'pass': continue
            if status == 'Fail' and pf != 'fail': continue
            filtered[key] = p
        self._populate_tree(filtered)
    
    def clear_filters(self):
        self.region_cb.set('All')
        self.status_cb.set('All')
        self._populate_tree()
    
    def on_tree_select(self, event):
        selected = self.tree.selection()
        self.sel_lbl.config(text=f"Selected: {len(selected)} of {len(self.tree.get_children())}")
    
    def select_all(self):
        self.tree.selection_set(self.tree.get_children())
        self.on_tree_select(None)
    
    def configure_template(self):
        current = self.config.get('builder_home_id_template', '{Subdivision1}_Lot{Lot1}')
        new_template = simpledialog.askstring("Template", f"Builder Home ID Template:\n\nAvailable: {{Subdivision1}}, {{Lot1}}, {{PermitNo1}}, {{RTIN}}", initialvalue=current, parent=self.root)
        if new_template:
            self.config.set('builder_home_id_template', new_template)
            messagebox.showinfo("Saved", f"Template: {new_template}")
    
    # ================================================================
    # EXPORT
    # ================================================================
    
    def preview_json(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Select", "Select projects first")
            return
        projects = [self.all_projects.get(k) for k in selected[:3]]
        version = self.version_cb.get()
        orientation = self.orientation_cb.get().split(' ')[0]
        data = self.json_gen.generate(projects, version, orientation)
        win = tk.Toplevel(self.root)
        win.title("JSON Preview")
        win.geometry("700x500")
        txt = tk.Text(win, font=('Consolas', 10), bg=self.theme.get('bg_alt'), fg=self.theme.get('fg'))
        txt.pack(fill='both', expand=True)
        txt.insert('1.0', json.dumps(data, indent=2))
    
    def generate_json(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Select", "Select projects to export")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not filepath: return
        projects = [self.all_projects.get(k) for k in selected]
        version = self.version_cb.get()
        orientation = self.orientation_cb.get().split(' ')[0]
        data = self.json_gen.generate(projects, version, orientation)
        with open(filepath, 'w') as f:
            json.dump(data, f, indent=2)
        self.status.config(text=f"Exported {len(selected)} projects to {filepath}")
        messagebox.showinfo("Export", f"Exported {len(selected)} projects")
    
    def export_rem_xml(self):
        if not self.all_projects:
            messagebox.showwarning("No Data", "Load data first")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML", "*.xml")])
        if not filepath: return
        try:
            REMFileHandler.export_to_rem_xml(list(self.all_projects.values()), filepath)
            self.status.config(text=f"Exported REM XML to {filepath}")
            messagebox.showinfo("Export", f"Exported {len(self.all_projects)} projects to REM XML")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
    
    def export_rem_csv(self):
        if not self.all_projects:
            messagebox.showwarning("No Data", "Load data first")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not filepath: return
        try:
            REMFileHandler.export_to_rem_csv(list(self.all_projects.values()), filepath)
            self.status.config(text=f"Exported REM CSV to {filepath}")
            messagebox.showinfo("Export", f"Exported {len(self.all_projects)} projects to REM CSV")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
    
    # ================================================================
    # VALIDATION
    # ================================================================
    
    def run_validation(self):
        if not self.all_projects:
            messagebox.showwarning("No Data", "Load data first")
            return
        self.validation_results.clear()
        self.val_tree.delete(*self.val_tree.get_children())
        errors = warnings = 0
        for key, p in self.all_projects.items():
            result = self.validator.validate_project(p)
            self.validation_results[key] = result
            status = '[OK] Valid' if result['is_valid'] else '[X] Invalid'
            tag = 'pass' if result['is_valid'] else 'fail'
            self.val_tree.insert('', 'end', iid=key, values=(key[:30], len(result['errors']), len(result['warnings']), status), tags=(tag,))
            errors += len(result['errors'])
            warnings += len(result['warnings'])
        valid = sum(1 for r in self.validation_results.values() if r['is_valid'])
        self.val_sum.config(text=f"Valid: {valid}/{len(self.all_projects)} | Errors: {errors} | Warnings: {warnings}")
    
    def show_val_details(self, event):
        selected = self.val_tree.selection()
        if not selected: return
        key = selected[0]
        result = self.validation_results.get(key, {})
        self.val_txt.delete('1.0', 'end')
        self.val_txt.insert('end', f"PROJECT: {key}\n{'='*50}\n\n")
        if result.get('errors'):
            self.val_txt.insert('end', "ERRORS:\n")
            for e in result['errors']:
                self.val_txt.insert('end', f"  [X] {e}\n")
        if result.get('warnings'):
            self.val_txt.insert('end', "\nWARNINGS:\n")
            for w in result['warnings']:
                self.val_txt.insert('end', f"  [!] {w}\n")
        if result.get('is_valid'):
            self.val_txt.insert('end', "\n[OK] Project is valid for export\n")
    
    # ================================================================
    # COMPLIANCE
    # ================================================================
    
    def run_compliance(self):
        if not self.all_projects:
            messagebox.showwarning("No Data", "Load data first")
            return
        self.compliance_results.clear()
        self.comp_tree.delete(*self.comp_tree.get_children())
        standard = ComplianceStandards.get_standard(self.std_cb.get())
        checker = ComplianceChecker(standard)
        pass_count = fail_count = warn_count = 0
        for key, p in self.all_projects.items():
            result = checker.check_project(p)
            self.compliance_results[key] = result
            tag = 'pass' if result['overall'] == 'PASS' else 'fail' if result['overall'] == 'FAIL' else 'warn'
            self.comp_tree.insert('', 'end', iid=key, values=(key[:30], result['pass_count'], result['fail_count'], result['warn_count'], result['overall']), tags=(tag,))
            if result['overall'] == 'PASS': pass_count += 1
            elif result['overall'] == 'FAIL': fail_count += 1
            else: warn_count += 1
        self.comp_sum.config(text=f"Pass: {pass_count} | Fail: {fail_count} | Warn: {warn_count}")
    
    def show_comp_details(self, event):
        selected = self.comp_tree.selection()
        if not selected: return
        key = selected[0]
        result = self.compliance_results.get(key, {})
        self.comp_txt.delete('1.0', 'end')
        self.comp_txt.insert('end', f"PROJECT: {key}\n{'='*50}\n\n")
        if result.get('footnotes_applied'):
            self.comp_txt.insert('end', "FOOTNOTES APPLIED:\n")
            for fn in result['footnotes_applied']:
                self.comp_txt.insert('end', f"  * {fn}\n")
            self.comp_txt.insert('end', "\n")
        self.comp_txt.insert('end', "COMPLIANCE CHECKS:\n\n")
        for check in result.get('checks', []):
            icon = '[OK]' if check['status'] == 'PASS' else '[X]' if check['status'] == 'FAIL' else '[!]'
            self.comp_txt.insert('end', f"{icon} {check['component']}\n")
            self.comp_txt.insert('end', f"   Value: {check['value']}\n")
            self.comp_txt.insert('end', f"   Requirement: {check['requirement']}\n\n")
        self.comp_txt.insert('end', f"\nOVERALL: {result.get('overall', 'N/A')}\n")
    
    # ================================================================
    # CHARTS
    # ================================================================
    
    def refresh_charts(self):
        if not HAS_MATPLOTLIB: return
        for w in self.chart_frame.winfo_children():
            w.destroy()
        if not self.all_projects:
            ttk.Label(self.chart_frame, text="Load data to see charts", font=('Arial', 12)).pack(pady=50)
            return
        chart_type = self.chart_type_cb.get()
        if chart_type == 'Overview':
            self._draw_overview_charts()
        elif chart_type == 'Duct Leakage':
            self._draw_duct_charts()
        elif chart_type == 'HVAC':
            self._draw_hvac_charts()
        elif chart_type == 'Static Pressure':
            self._draw_pressure_charts()
        elif chart_type == 'By Region':
            self._draw_region_charts()
    
    def _draw_overview_charts(self):
        fig = Figure(figsize=(12, 8), dpi=100, facecolor=self.theme.get('bg'))
        projects = list(self.all_projects.values())
        
        # Pass/Fail pie
        ax1 = fig.add_subplot(2, 2, 1, facecolor=self.theme.get('bg_alt'))
        pf = [str(p.get('PassFail1', '')).lower() for p in projects]
        passes = pf.count('pass')
        fails = pf.count('fail')
        other = len(pf) - passes - fails
        if passes or fails or other:
            ax1.pie([passes, fails, other], labels=['Pass', 'Fail', 'Other'], autopct='%1.0f%%', colors=['#28a745', '#dc3545', '#6c757d'])
        ax1.set_title('Pass/Fail Distribution', color=self.theme.get('fg'))
        
        # TDL distribution
        ax2 = fig.add_subplot(2, 2, 2, facecolor=self.theme.get('bg_alt'))
        tdl_vals = [p.get('TDLCFM') for p in projects if p.get('TDLCFM') is not None]
        if tdl_vals:
            ax2.hist(tdl_vals, bins=20, color=self.theme.get('accent'), edgecolor='white')
        ax2.set_xlabel('TDL CFM25', color=self.theme.get('fg'))
        ax2.set_ylabel('Count', color=self.theme.get('fg'))
        ax2.set_title('Total Duct Leakage Distribution', color=self.theme.get('fg'))
        ax2.tick_params(colors=self.theme.get('fg'))
        
        # Rating types
        ax3 = fig.add_subplot(2, 2, 3, facecolor=self.theme.get('bg_alt'))
        ratings = [RatingType.determine(p) for p in projects]
        confirmed = ratings.count('Confirmed')
        projected = ratings.count('Projected')
        ax3.bar(['Confirmed', 'Projected'], [confirmed, projected], color=[self.theme.get('success'), self.theme.get('warning')])
        ax3.set_title('Rating Types', color=self.theme.get('fg'))
        ax3.tick_params(colors=self.theme.get('fg'))
        
        # By region
        ax4 = fig.add_subplot(2, 2, 4, facecolor=self.theme.get('bg_alt'))
        regions = {}
        for p in projects:
            r = str(p.get('Region', 'Unknown'))[:15]
            regions[r] = regions.get(r, 0) + 1
        if regions:
            top = sorted(regions.items(), key=lambda x: -x[1])[:8]
            ax4.barh([t[0] for t in top], [t[1] for t in top], color=self.theme.get('accent'))
        ax4.set_title('Projects by Region', color=self.theme.get('fg'))
        ax4.tick_params(colors=self.theme.get('fg'))
        
        fig.tight_layout(pad=3.0)
        canvas = FigureCanvasTkAgg(fig, self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    
    def _draw_duct_charts(self):
        fig = Figure(figsize=(12, 8), dpi=100, facecolor=self.theme.get('bg'))
        projects = list(self.all_projects.values())
        
        # TDL vs LTO scatter
        ax1 = fig.add_subplot(2, 2, 1, facecolor=self.theme.get('bg_alt'))
        tdl = [p.get('TDLCFM') for p in projects if p.get('TDLCFM') is not None and p.get('LTOCFM') is not None]
        lto = [p.get('LTOCFM') for p in projects if p.get('TDLCFM') is not None and p.get('LTOCFM') is not None]
        if tdl and lto:
            ax1.scatter(tdl, lto, alpha=0.6, c=self.theme.get('accent'))
            ax1.plot([0, max(tdl)], [0, max(tdl)], 'r--', alpha=0.5, label='1:1')
        ax1.set_xlabel('TDL CFM25', color=self.theme.get('fg'))
        ax1.set_ylabel('LTO CFM25', color=self.theme.get('fg'))
        ax1.set_title('TDL vs LTO', color=self.theme.get('fg'))
        ax1.tick_params(colors=self.theme.get('fg'))
        
        # TDL rate distribution
        ax2 = fig.add_subplot(2, 2, 2, facecolor=self.theme.get('bg_alt'))
        rates = []
        for p in projects:
            if p.get('TDLCFM') is not None and p.get('Living') and p['Living'] > 0:
                rates.append((p['TDLCFM'] / p['Living']) * 100)
        if rates:
            ax2.hist(rates, bins=20, color=self.theme.get('accent'), edgecolor='white')
            ax2.axvline(8, color='red', linestyle='--', label='Limit (8)')
            ax2.axvline(12, color='orange', linestyle='--', label='Fn41 (12)')
        ax2.set_xlabel('CFM25/100sqft', color=self.theme.get('fg'))
        ax2.set_title('TDL Rate Distribution', color=self.theme.get('fg'))
        ax2.tick_params(colors=self.theme.get('fg'))
        ax2.legend()
        
        # LTO rate distribution
        ax3 = fig.add_subplot(2, 2, 3, facecolor=self.theme.get('bg_alt'))
        lto_rates = []
        for p in projects:
            if p.get('LTOCFM') is not None and p.get('Living') and p['Living'] > 0:
                lto_rates.append((p['LTOCFM'] / p['Living']) * 100)
        if lto_rates:
            ax3.hist(lto_rates, bins=20, color='#17a2b8', edgecolor='white')
            ax3.axvline(4, color='red', linestyle='--', label='Limit (4)')
        ax3.set_xlabel('CFM25/100sqft', color=self.theme.get('fg'))
        ax3.set_title('LTO Rate Distribution', color=self.theme.get('fg'))
        ax3.tick_params(colors=self.theme.get('fg'))
        ax3.legend()
        
        # TDL vs Sqft
        ax4 = fig.add_subplot(2, 2, 4, facecolor=self.theme.get('bg_alt'))
        sqft = [p.get('Living') for p in projects if p.get('TDLCFM') is not None and p.get('Living')]
        tdl_v = [p.get('TDLCFM') for p in projects if p.get('TDLCFM') is not None and p.get('Living')]
        if sqft and tdl_v:
            ax4.scatter(sqft, tdl_v, alpha=0.5, c=self.theme.get('accent'))
        ax4.set_xlabel('Living Sqft', color=self.theme.get('fg'))
        ax4.set_ylabel('TDL CFM25', color=self.theme.get('fg'))
        ax4.set_title('TDL vs Living Area', color=self.theme.get('fg'))
        ax4.tick_params(colors=self.theme.get('fg'))
        
        fig.tight_layout(pad=3.0)
        canvas = FigureCanvasTkAgg(fig, self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    
    def _draw_hvac_charts(self):
        fig = Figure(figsize=(12, 8), dpi=100, facecolor=self.theme.get('bg'))
        projects = list(self.all_projects.values())
        
        # CFM/Ton histogram
        ax1 = fig.add_subplot(2, 2, 1, facecolor=self.theme.get('bg_alt'))
        cpt = []
        for p in projects:
            if p.get('MeasuredCFM') and p.get('Tonnage') and p['Tonnage'] > 0:
                cpt.append(p['MeasuredCFM'] / p['Tonnage'])
        if cpt:
            ax1.hist(cpt, bins=20, color=self.theme.get('accent'), edgecolor='white')
            ax1.axvline(350, color='red', linestyle='--')
            ax1.axvline(450, color='red', linestyle='--')
        ax1.set_xlabel('CFM/Ton', color=self.theme.get('fg'))
        ax1.set_title('Airflow per Ton', color=self.theme.get('fg'))
        ax1.tick_params(colors=self.theme.get('fg'))
        
        # Charge variance
        ax2 = fig.add_subplot(2, 2, 2, facecolor=self.theme.get('bg_alt'))
        charges = [p.get('Charge') for p in projects if p.get('Charge') is not None]
        if charges:
            ax2.hist(charges, bins=20, color='#17a2b8', edgecolor='white')
            ax2.axvline(-0.05, color='red', linestyle='--')
            ax2.axvline(0.05, color='red', linestyle='--')
        ax2.set_xlabel('Charge Variance', color=self.theme.get('fg'))
        ax2.set_title('Refrigerant Charge', color=self.theme.get('fg'))
        ax2.tick_params(colors=self.theme.get('fg'))
        
        # Tonnage vs Sqft
        ax3 = fig.add_subplot(2, 2, 3, facecolor=self.theme.get('bg_alt'))
        sqft = [p.get('Living') for p in projects if p.get('Tonnage') and p.get('Living')]
        tons = [p.get('Tonnage') for p in projects if p.get('Tonnage') and p.get('Living')]
        if sqft and tons:
            ax3.scatter(sqft, tons, alpha=0.5, c=self.theme.get('accent'))
        ax3.set_xlabel('Living Sqft', color=self.theme.get('fg'))
        ax3.set_ylabel('Tonnage', color=self.theme.get('fg'))
        ax3.set_title('Tonnage vs Living Area', color=self.theme.get('fg'))
        ax3.tick_params(colors=self.theme.get('fg'))
        
        # Wattage
        ax4 = fig.add_subplot(2, 2, 4, facecolor=self.theme.get('bg_alt'))
        watts = [p.get('MeasuredWattage') for p in projects if p.get('MeasuredWattage') is not None]
        if watts:
            ax4.hist(watts, bins=20, color='#ffc107', edgecolor='white')
        ax4.set_xlabel('Watts', color=self.theme.get('fg'))
        ax4.set_title('Measured Wattage', color=self.theme.get('fg'))
        ax4.tick_params(colors=self.theme.get('fg'))
        
        fig.tight_layout(pad=3.0)
        canvas = FigureCanvasTkAgg(fig, self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    
    def _draw_pressure_charts(self):
        fig = Figure(figsize=(12, 8), dpi=100, facecolor=self.theme.get('bg'))
        projects = list(self.all_projects.values())
        
        # Return vs Supply scatter
        ax1 = fig.add_subplot(2, 2, 1, facecolor=self.theme.get('bg_alt'))
        ret = [abs(p.get('ReturnIWC', 0)) for p in projects if p.get('ReturnIWC') is not None and p.get('SupplyIWC') is not None]
        sup = [abs(p.get('SupplyIWC', 0)) for p in projects if p.get('ReturnIWC') is not None and p.get('SupplyIWC') is not None]
        if ret and sup:
            ax1.scatter(ret, sup, alpha=0.6, c=self.theme.get('accent'))
            ax1.axvline(0.20, color='red', linestyle='--', alpha=0.7)
            ax1.axhline(0.25, color='red', linestyle='--', alpha=0.7)
        ax1.set_xlabel('Return IWC', color=self.theme.get('fg'))
        ax1.set_ylabel('Supply IWC', color=self.theme.get('fg'))
        ax1.set_title('Return vs Supply Static', color=self.theme.get('fg'))
        ax1.tick_params(colors=self.theme.get('fg'))
        
        # Return distribution
        ax2 = fig.add_subplot(2, 2, 2, facecolor=self.theme.get('bg_alt'))
        ret_vals = [abs(p.get('ReturnIWC', 0)) for p in projects if p.get('ReturnIWC') is not None]
        if ret_vals:
            ax2.hist(ret_vals, bins=20, color='#28a745', edgecolor='white')
            ax2.axvline(0.20, color='red', linestyle='--', label='Max 0.20')
        ax2.set_xlabel('IWC', color=self.theme.get('fg'))
        ax2.set_title('Return Static Pressure', color=self.theme.get('fg'))
        ax2.tick_params(colors=self.theme.get('fg'))
        ax2.legend()
        
        # Supply distribution
        ax3 = fig.add_subplot(2, 2, 3, facecolor=self.theme.get('bg_alt'))
        sup_vals = [abs(p.get('SupplyIWC', 0)) for p in projects if p.get('SupplyIWC') is not None]
        if sup_vals:
            ax3.hist(sup_vals, bins=20, color='#17a2b8', edgecolor='white')
            ax3.axvline(0.25, color='red', linestyle='--', label='Max 0.25')
        ax3.set_xlabel('IWC', color=self.theme.get('fg'))
        ax3.set_title('Supply Static Pressure', color=self.theme.get('fg'))
        ax3.tick_params(colors=self.theme.get('fg'))
        ax3.legend()
        
        # Total ESP
        ax4 = fig.add_subplot(2, 2, 4, facecolor=self.theme.get('bg_alt'))
        esp = [abs(p.get('ReturnIWC', 0)) + abs(p.get('SupplyIWC', 0)) for p in projects if p.get('ReturnIWC') is not None and p.get('SupplyIWC') is not None]
        if esp:
            ax4.hist(esp, bins=20, color='#6f42c1', edgecolor='white')
            ax4.axvline(0.45, color='red', linestyle='--', label='Max 0.45')
        ax4.set_xlabel('Total ESP (IWC)', color=self.theme.get('fg'))
        ax4.set_title('Total External Static Pressure', color=self.theme.get('fg'))
        ax4.tick_params(colors=self.theme.get('fg'))
        ax4.legend()
        
        fig.tight_layout(pad=3.0)
        canvas = FigureCanvasTkAgg(fig, self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    
    def _draw_region_charts(self):
        fig = Figure(figsize=(12, 8), dpi=100, facecolor=self.theme.get('bg'))
        projects = list(self.all_projects.values())
        
        regions = {}
        for p in projects:
            r = str(p.get('Region', 'Unknown'))[:20]
            if r not in regions:
                regions[r] = {'count': 0, 'pass': 0, 'tdl_sum': 0, 'tdl_count': 0}
            regions[r]['count'] += 1
            if str(p.get('PassFail1', '')).lower() == 'pass':
                regions[r]['pass'] += 1
            if p.get('TDLCFM') is not None:
                regions[r]['tdl_sum'] += p['TDLCFM']
                regions[r]['tdl_count'] += 1
        
        top = sorted(regions.items(), key=lambda x: -x[1]['count'])[:10]
        names = [t[0] for t in top]
        counts = [t[1]['count'] for t in top]
        pass_rates = [t[1]['pass'] / t[1]['count'] * 100 if t[1]['count'] > 0 else 0 for t in top]
        avg_tdl = [t[1]['tdl_sum'] / t[1]['tdl_count'] if t[1]['tdl_count'] > 0 else 0 for t in top]
        
        # Count by region
        ax1 = fig.add_subplot(2, 2, 1, facecolor=self.theme.get('bg_alt'))
        ax1.barh(names, counts, color=self.theme.get('accent'))
        ax1.set_xlabel('Project Count', color=self.theme.get('fg'))
        ax1.set_title('Projects by Region', color=self.theme.get('fg'))
        ax1.tick_params(colors=self.theme.get('fg'))
        
        # Pass rate by region
        ax2 = fig.add_subplot(2, 2, 2, facecolor=self.theme.get('bg_alt'))
        colors = ['#28a745' if r >= 90 else '#ffc107' if r >= 70 else '#dc3545' for r in pass_rates]
        ax2.barh(names, pass_rates, color=colors)
        ax2.set_xlabel('Pass Rate (%)', color=self.theme.get('fg'))
        ax2.set_title('Pass Rate by Region', color=self.theme.get('fg'))
        ax2.tick_params(colors=self.theme.get('fg'))
        ax2.set_xlim(0, 100)
        
        # Avg TDL by region
        ax3 = fig.add_subplot(2, 2, 3, facecolor=self.theme.get('bg_alt'))
        ax3.barh(names, avg_tdl, color='#17a2b8')
        ax3.set_xlabel('Avg TDL CFM25', color=self.theme.get('fg'))
        ax3.set_title('Avg TDL by Region', color=self.theme.get('fg'))
        ax3.tick_params(colors=self.theme.get('fg'))
        
        # Pie of top regions
        ax4 = fig.add_subplot(2, 2, 4, facecolor=self.theme.get('bg_alt'))
        if counts:
            ax4.pie(counts[:6], labels=names[:6], autopct='%1.0f%%')
        ax4.set_title('Top Regions Share', color=self.theme.get('fg'))
        
        fig.tight_layout(pad=3.0)
        canvas = FigureCanvasTkAgg(fig, self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    
    # ================================================================
    # CALCULATORS
    # ================================================================
    
    def calc_duct_limits(self):
        try:
            sqft = float(self.dl_sqft.get())
            returns = int(self.dl_returns.get())
            measured = float(self.dl_measured.get())
            std_limit = max((sqft / 100) * 8, 80)
            fn41_limit = max((sqft / 100) * 12, 120)
            use_fn41 = returns >= 3
            applicable = fn41_limit if use_fn41 else std_limit
            status = "[OK] PASS" if measured <= applicable else "[X] FAIL"
            result = f"Living: {sqft:.0f} sqft\nReturns: {returns}\n\n"
            result += f"Standard Limit: {std_limit:.0f} CFM25\nFn41 Limit: {fn41_limit:.0f} CFM25\n\n"
            result += f"Applicable: {applicable:.0f} CFM25\nMeasured: {measured:.0f} CFM25\n\n{status}"
            self.dl_res.config(text=result)
        except:
            messagebox.showerror("Error", "Invalid input")
    
    def calc_cfm_per_ton(self):
        try:
            cfm = float(self.ct_cfm.get())
            tons = float(self.ct_ton.get())
            result = self.calc.cfm_per_ton(cfm, tons)
            status = "[OK] GOOD" if 350 <= result <= 450 else "[!] CHECK"
            self.ct_res.config(text=f"{result:.0f} CFM/Ton\n{status}\n(Target: 350-450)")
        except:
            messagebox.showerror("Error", "Invalid input")
    
    def calc_static_pressure(self):
        try:
            ret = float(self.sp_ret.get())
            sup = float(self.sp_sup.get())
            total = self.calc.total_external_sp(ret, sup)
            ret_status = "[OK]" if ret <= 0.20 else "[!]"
            sup_status = "[OK]" if sup <= 0.25 else "[!]"
            result = f"Return: {ret:.3f} IWC {ret_status} (max 0.20)\n"
            result += f"Supply: {sup:.3f} IWC {sup_status} (max 0.25)\n\n"
            result += f"Total ESP: {total:.3f} IWC"
            self.sp_res.config(text=result)
        except:
            messagebox.showerror("Error", "Invalid input")
    
    def calc_ach50(self):
        try:
            bd = float(self.ach_bd.get())
            sqft = float(self.ach_sqft.get())
            height = float(self.ach_height.get())
            ach50 = self.calc.ach50(bd, sqft, height)
            natural = self.calc.natural_ach(ach50)
            rating = "Tight" if ach50 < 5 else "Average" if ach50 < 7 else "Leaky"
            result = f"BD: {bd:.0f} CFM50\nVolume: {sqft * height:.0f} cuft\n\n"
            result += f"ACH50: {ach50:.2f}\nNatural ACH: {natural:.3f}\n\nRating: {rating}"
            self.ach_res.config(text=result)
        except:
            messagebox.showerror("Error", "Invalid input")
    
    def calc_ventilation(self):
        try:
            sqft = float(self.vent_sqft.get())
            br = int(self.vent_br.get())
            cfm = self.calc.required_ventilation_cfm(sqft, br)
            result = f"Living: {sqft:.0f} sqft\nBedrooms: {br}\n\n"
            result += f"ASHRAE 62.2:\n0.01  {sqft:.0f} + 7.5  ({br}+1)\n\n"
            result += f"Required: {cfm:.0f} CFM"
            self.vent_res.config(text=result)
        except:
            messagebox.showerror("Error", "Invalid input")
    
    # ================================================================
    # HELP
    # ================================================================
    
    def show_about(self):
        messagebox.showinfo("About",
            "DSLD Homes - Ekotrope Sync v9\n\n"
            "ENERGY STAR 3.2 Compliant\n"
            "REM/Rate Integration\n\n"
            "Features:\n"
            "* ENERGY STAR 3.2 compliance\n"
            "* Footnotes 40, 41, 44, 45\n"
            "* REM/Rate import/export\n"
            "* Extended calculators\n"
            "* Enhanced charts\n"
            "* Dark/Light mode\n\n"
            " 2026 DSLD Homes"
        )
    
    def show_compliance_ref(self):
        win = tk.Toplevel(self.root)
        win.title("ENERGY STAR 3.2 Reference")
        win.geometry("550x450")
        win.configure(bg=self.theme.get('bg'))
        txt = tk.Text(win, font=('Consolas', 10), wrap='word', bg=self.theme.get('bg_alt'), fg=self.theme.get('fg'))
        txt.pack(fill='both', expand=True, padx=10, pady=10)
        ref = """ENERGY STAR 3.2 COMPLIANCE REFERENCE
=====================================

6.4.2 - TOTAL DUCT LEAKAGE (FINAL)
Standard: <=8 CFM25/100sf or <=80 CFM25
Footnote 41 (3+ returns): <=12 CFM25/100sf or <=120 CFM25

6.5 - LEAKAGE TO OUTSIDE
Standard: <=4 CFM25/100sf or <=40 CFM25

FOOTNOTE 45 - LTO WAIVER
If TDL <=4 CFM25/100sf or <=40 CFM25

5b.2 - STATIC PRESSURE
Return: <=0.20 IWC
Supply: <=0.25 IWC

5a.3 - CHARGE: +/-0.05

5a.1 - AIRFLOW: 350-450 CFM/ton

8.2 - BATH FAN: >=50 CFM intermittent

RATING TYPES
Confirmed: Final + Pass
Projected: Otherwise
"""
        txt.insert('1.0', ref)
        txt.config(state='disabled')
    
    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = EkotropeSyncApp()
    app.run()
