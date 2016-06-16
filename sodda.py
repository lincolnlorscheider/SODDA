import csv
import tkFileDialog
import os
import os.path
import re

__author__ = 'Lincoln Lorscheider'



class Report(object):
    """ Report class opens a .csv file path and reads it into the Report.data attribute, which is a csv.reader instance
     unless the file was a .txt file, in which case it's a list"""
    def __init__(self, report_path=None):
        """A class that opens csv reports and reads them into memory"""
        if report_path is None:
            report_path = tkFileDialog.askopenfilename(title="Select the Application Subpoint Report you wish to analyze")
        self.report_path = report_path
        self.report_file_extension = os.path.split(self.report_path)[1]
        if self.report_file_extension.endswith(".txt"):
            self.data = list(open(self.report_path))
        else:
            self.load_csv(report_path)

    def load_csv(self, filepath=None):
        if filepath is not None:
            if os.path.exists(filepath):
                self.report_path = filepath
            else:
                return InvalidFilePath
        reader = csv.reader(open(self.report_path,'rb'))
        self.data = reader


class SubpointReport(Report):
    """Inherits from Report.  Populates and Exposes a list of TEC objects from the application subpoint report it takes as its input."""
    def __init__(self, report_path=None):
        super(SubpointReport, self).__init__(report_path)
        self.TECs = []
        self._headers = []
        self.failures = []
        self.create_TECs()
        self.analyze_tecs()
        self.dump_analysis()

    def create_TECs(self):
        """populates self.TECs with the devices defined in the report"""
        tec = None
        for row in self.data:
            if len(row) > 1:
                if row[0] == "TEC System Name:":
                    if tec is not None:
                        self.TECs.append(tec)
                    tec = TEC(row[1])
                    tec.descriptor = row[4]
                else:
                    subpoint = Point()
                    subpoint.name = row[1].split(":")[1].strip()
                    if subpoint.name == "ADDRESS":
                        subpoint.address = 1
                    else:
                        subpoint.address = int(row[0].strip())
                    subpoint.device = tec.name
                    if subpoint.name == "APPLICATION":
                        subpoint.value = int(float(row[4]))
                        tec.application = int(float(row[4]))
                    else:
                        subpoint.value = row[4]
                    subpoint.units = row[5]
                    subpoint.status, subpoint.priority = row[6].split("    ")
                    subpoint.priority = subpoint.priority.strip()
                    subpoint.status = subpoint.status.strip()
                    tec.subpoints.append(subpoint)
            else:
                if len(row[0].split("**********"))>1:
                    self.TECs.append(tec)
                    break

    def analyze_tecs(self):
        self.failures = []
        for tec in self.TECs:
            self._headers.append(tec.name)
            for index, failure in enumerate(tec.analyze()):
                try:
                    self.failures[index][tec.name] = failure
                except IndexError:
                    self.failures.append({tec.name: failure})

    def dump_analysis(self):
        outfile = tkFileDialog.asksaveasfile()
        writer = csv.writer(outfile, lineterminator='\n')
        for items in self.failures:
            for k,v in items.items():
                writer.writerow([k,v])


class TEC(object):

    def __init__(self, name):
        self.name = name
        self.descriptor =''
        self.application = ''
        self.subpoints=[]
        self.status = ''
        self.failures = []

    def __repr__(self):
        return self.name

    def analyze(self):
        if str(self.application).endswith("90") or str(self.application).endswith("91") or str(self.application).endswith("92"):
            self.failures.append("This device is in slave mode")
        if self.is_failed() is False:
            self.check_sensors()
            self.check_dampers()
            self.compare_temp_to_setpoint()
            self.compare_flow_to_setpoint()
            self.sanity_check()
        return self.failures

    def compare_temp_to_setpoint(self, differential=5): #Todo Finish this
        room_temp = None
        temp_stpt = None
        failure = []
        for point in self.subpoints:
            if point.name == "CTL TEMP":
                room_temp = float(point.value)
            if point.name == "CTL STPT":
                temp_stpt = float(point.value)
        if room_temp is not None and temp_stpt is not None:
            if abs(temp_stpt - room_temp) >= differential:
                failure.append("Temperature Control Failure, CTL TEMP is currently {} degrees from CTL STPT".format(temp_stpt-room_temp))
                self.failures.extend(failure)
                return True
        return False

    def compare_flow_to_setpoint(self, differential=25):
        failure = []
        air_volume = None
        flow_setpoint_pct = None
        occupance = None
        min_flow = None
        max_flow = None
        flow_pct = None
        for point in self.subpoints:
            if point.name == "FLOW":
                flow_pct = float(point.value)
            if point.name == "FLOW STPT":
                flow_setpoint_pct = float(point.value)
        if flow_setpoint_pct is not None and flow_pct is not None:
            delta = flow_setpoint_pct - flow_pct
            if abs(delta) > differential:
                failure.append("Flow Control Failure, Flow is currently {} % from Flow Setpoint".format(delta))
                self.failures.extend(failure)
                return True
        return False

    def check_sensors(self):
        failure = []
        for point in self.subpoints:
            if "AIR VOL" in point.name:
                if point.status == Point.failed:
                    failure.append("The {} sensor is Failed".format(point.name))
            if point.name == "ROOM TEMP":
                if point.status == Point.failed:
                    failure.append("The {} sensor is Failed".format(point.name))
        if len(failure) > 0:
            self.failures.extend(failure)
            return True
        return False

    def check_dampers(self):
        """ This test compares air volume against damper position.  UME is standard VAV, GEX, and SUP are exhaust and supply."""
        failure = []
        airflows = {}
        damper_positions = {}
        for point in self.subpoints:
            if "AIR VOL" in point.name:
                airflows[point.name.replace("AIR VOL", "")] = float(point.value)
            if "DMP" in point.name and (point.name.endswith("CMD") or point.name.endswith("COMD")):
                if point.name == "DMPR COMD":
                    damper_positions["UME"] = float(point.value)
                else:
                    damper_positions[point.name.replace("DMP CMD", "")] = float(point.value)
        for duct, airflow in airflows.items():
            dmpr_position = damper_positions[duct]
            if dmpr_position == 0 and airflow > 10:
                failure.append("Slipped Damper - Airflow greater than 10 cfm across closed damper")
            if dmpr_position > 85:
                    if airflow < 150:
                        failure.append("Slipped Damper - Airflow less than than 150 cfm across fully open damper")
                    else:
                        failure.append("Potential Starved Zone - Damper {}% open with {} cfm of airflow".format(dmpr_position, airflow))
        if len(failure) > 0:
            self.failures.extend(failure)
            return True
        return False

    def sanity_check(self):
        failure = []
        occ_flow = None
        unocc_flow = None
        for point in self.subpoints:
            if point.name == "OCC FLOW":
                occ_flow = float(point.value)
            if point.name == "UNOCC FLOW":
                unocc_flow = float(point.value)
            if point.name == "CTL STPT" or "LOOPOUT" in point.name:
                if point.priority != "NONE":
                    failure.append("{} Not in Automatic - Automatic operation of this point is essential for proper operation".format(point.name))
        if occ_flow is not None and unocc_flow is not None:
            if occ_flow == 0:
                failure.append("Occ flow setpoint - The occupied mode has an airflow setpoint of zero")
            if unocc_flow > occ_flow:
                failure.append("Unocc flow setpoint - The unoccupied mode has a higher airflow setpoint than occupied")
        if len(failure) != 0:
            self.failures.extend(failure)
            return True
        return False

    def is_failed(self):
        failure = []
        for point in self.subpoints:
            if point.name == "APPLICATION" and point.status == Point.failed:
                failure.append("NOT COMMUNICATING - {} controller may be broken".format(point.device))
                self.failures.extend(failure)
                return True
        return False


class Point(object):

    failed = "*F*"
    normal = "-N-"

    def __init__(self):
        self.device = ''
        self.address = ''
        self.name = '' #used for system name on all but PDS points
        self.value = ''
        self.units = ''
        self.status = ''
        self.priority = ''
        self.descriptor = ''
        self.system_name = ''
        self.wire_resistance = ''
        self.totalization = ''
        self.standard_alarms = ''
        self.special_mode_5 = ''
        self.special_mode_4= ''
        self.special_mode_3 = ''
        self.special_mode_2 = ''
        self.slope = ''
        self.setpoint_value_5 = ''
        self.setpoint_value_4 = ''
        self.setpoint_value_3 = ''
        self.setpoint_value_2 = ''
        self.setpoint_value_1 = ''
        self.setpoint_value_0 = ''
        self.setpoint_name_5 = ''
        self.setpoint_name_4 = ''
        self.setpoint_name_3 = ''
        self.setpoint_name_2 = ''
        self.setpoint_name_1 = ''
        self.setpoint_name_0 = ''
        self.sensor_type = ''
        self.reno = ''
        self.priority = ''
        self.popup = ''
        self.point_type = ''
        self.point_memo = ''
        self.point_address = ''
        self.panel_name = ''
        self.out_of_service = ''
        self.normal_ack_enabled = ''
        self.night_mode_0 = ''
        self.mode_delay = ''
        self.low_alarm_limit = ''
        self.level_delay = ''
        self.intercept = ''
        self.initial_value = ''
        self.initial_priority = ''
        self.informational_text = ''
        self.high_alarm_limit = ''
        self.graphic_name = ''
        self.enhanced_alarms = ''
        self.enhanced_alarm_mode_point = ''
        self.engineering_units = ''
        self.differential = ''
        self.descriptor = ''
        self.day_mode_1 = ''
        self.cov_limit = ''
        self.classification = ''
        self.analog_representation = ''
        self.alarmable = ''
        self.alarm_message = ''
        self.alarm_destinations = ''
        self.aim = ''
        self.address_type = ''
        self.actuator_type = ''
        self.number_of_decimal_places = ''

    def __repr__(self):
        return str([self.address, self.device, self.name, self.value, self.units, self.status, self.priority])

    def is_out_of_auto(self):
        if self.priority !="NONE" and self.priority != "OVRD" and not(self.name.endswith("STPT")):
            return True
        else:
            return False

    def is_out_of_normal(self):
        if self.status != Point.normal:
            return True
        else:
            return False


class InvalidFilePath(Exception):
    def __init__(self, message):
        self.message = message

    def __repr__(self):
        return self.message


class PanelPPCLReport(Report):
    def __init__(self, report_path=None):
        super(PanelPPCLReport, self).__init__(report_path)
        if self.report_file_extension.endswith(".csv"):
            raise InvalidFilePath("This application only accepts Panel PPCL Reports in .txt format")
        else:
            self.convert()

    def convert(self, file_as_list=None):
        prog_check_regex = '^Program Name:(.+)'
        if file_as_list is not None:
            self.data = file_as_list
        pcl_files = {}
        for line in self.data:
            prog_check = re.split(prog_check_regex, line)
            if len(prog_check) > 1:
                program_name = prog_check[1].strip()
            fractline = re.split('^\s{12,13}(\S.*)', line)
            if len(fractline) > 1:
                pcl_files[program_name][linenumber] = "".join([pcl_files[program_name][linenumber], fractline[1]])
            newline = re.split("^[EUTD]{1,2}\s{3,6}([0-9]{1,5})\s+([^-]\S.*)", line)
            if len(newline) > 1:
                linenumber = int(newline[1])
                code = newline[2]
                try:
                    pcl_files[program_name][linenumber] = code
                except KeyError:
                    pcl_files[program_name] = {linenumber: code}
        for key, value in pcl_files.items():
            buffer_list = []
            for linenumber in sorted(value.keys()):
                buffer_list.append("\t".join([str(linenumber), value[linenumber]+"\n"]))
            newfile = open(key+".pcl", mode='w+')
            newfile.writelines(buffer_list)


class PanelPointLogReport(Report):
    def __init__(self, report_path=None):
        super(PanelPointLogReport, self).__init__(report_path)
        self.point_list = []
        self.analysis = {}
        self.build_points()
        self._analyze()

    def build_points(self):
        for line in self.data:
            if len(line[0]) < 100:
                if line[0] != "":
                    try:
                        new_point = self._map_normal(line)
                    except:
                        new_point = Point()
                        new_point.name = line[0]
                else:
                    new_point = self._map_abberation(line, new_point)
            if new_point.priority != "":
                self.point_list.append(new_point)

    def _map_normal(self, line):
        point = Point()
        point.name = line[0]
        point.address = line[2]
        point.descriptor = line[3]
        point.value = line[4]
        point.units = line[5]
        point.status, point.priority = line[6].split("  ", 1)
        point.status = point.status.strip()
        point.priority = point.priority.strip()
        return point

    def _map_abberation(self, line, point):
        point.address = line[1]
        point.descriptor = line[2]
        point.value = line[3]
        point.units = line[4]
        point.status, point.priority = line[5].split("  ", 1)
        point.status = point.status.strip()
        point.priority = point.priority.strip()
        return point

    def _analyze(self):
        for point in self.point_list:
            if point.priority !="NONE" and point.priority != "OVRD" and not(point.name.endswith("STPT")):
                if point.priority == "":
                    print point.name
                try:
                    self.analysis['Not in Auto'].append(point.name)
                except KeyError:
                    self.analysis['Not in Auto'] = []
                    self.analysis['Not in Auto'].append(point.name)
            if point.status != Point.normal:
                try:
                    self.analysis['Not in Normal'].append(point.name)
                except KeyError:
                    self.analysis['Not in Normal'] = []
                    self.analysis['Not in Normal'].append(point.name)


class PointDataSorter(Report):
    def __init__(self, report_path=None):
        self.data = []
        super(PointDataSorter, self).__init__(report_path)
        self.point_list = []
        self.analysis = {}
        self.build_points()
        self.analyze()
        print self.analysis

    def load_csv(self, filepath=None):
        if filepath is not None:
            if os.path.exists(filepath):
                self.report_path = filepath
            else:
                return InvalidFilePath
        infile = open(self.report_path,'rb')
        infile.seek(0)
        next(infile)
        reader = csv.DictReader(infile)
        for row in reader:
            self.data.append(row)
        self.data = self.data[:-1]

    def build_points(self):
        for row in self.data:
            point = Point()
            point.device = self._keyerr_as_emptystring(row, 'Panel Name')
            point.address = self._keyerr_as_emptystring(row, 'Point Address')
            point.name = self._keyerr_as_emptystring(row, 'Name') #used for system name on all but PDS points
            point.system_name = self._keyerr_as_emptystring(row, 'System Name')
            point.wire_resistance = self._keyerr_as_emptystring(row, 'Wire Resistance')
            point.totalization = self._keyerr_as_emptystring(row, 'Totalization')
            point.standard_alarms = self._keyerr_as_emptystring(row, 'Standard Alarms')
            point.special_mode_5 = self._keyerr_as_emptystring(row, 'Special Mode (5)')
            point.special_mode_4 = self._keyerr_as_emptystring(row, 'Special Mode (4)')
            point.special_mode_3 = self._keyerr_as_emptystring(row, 'Special Mode (3)')
            point.special_mode_2 = self._keyerr_as_emptystring(row, 'Special Mode (2)')
            point.slope = self._keyerr_as_emptystring(row, 'Slope')
            point.setpoint_value_5 = self._keyerr_as_emptystring(row, 'Setpoint Value(5)')
            point.setpoint_value_4 = self._keyerr_as_emptystring(row, 'Setpoint Value(4)')
            point.setpoint_value_3 = self._keyerr_as_emptystring(row, 'Setpoint Value(3)')
            point.setpoint_value_2 = self._keyerr_as_emptystring(row, 'Setpoint Value(2)')
            point.setpoint_value_1 = self._keyerr_as_emptystring(row, 'Setpoint Value(1)')
            point.setpoint_value_0 = self._keyerr_as_emptystring(row, 'Setpoint Value(0)')
            point.setpoint_name_5 = self._keyerr_as_emptystring(row, 'Setpoint Name(5)')
            point.setpoint_name_4 = self._keyerr_as_emptystring(row, 'Setpoint Name(4)')
            point.setpoint_name_3 = self._keyerr_as_emptystring(row, 'Setpoint Name(3)')
            point.setpoint_name_2 = self._keyerr_as_emptystring(row, 'Setpoint Name(2)')
            point.setpoint_name_1 = self._keyerr_as_emptystring(row, 'Setpoint Name(1)')
            point.setpoint_name_0 = self._keyerr_as_emptystring(row, 'Setpoint Name(0)')
            point.sensor_type = self._keyerr_as_emptystring(row, 'Sensor Type')
            point.reno = self._keyerr_as_emptystring(row, 'RENO')
            point.priority = self._keyerr_as_emptystring(row, 'Priority')
            point.popup = self._keyerr_as_emptystring(row, 'Popup')
            point.point_type = self._keyerr_as_emptystring(row, 'Point Type')
            point.point_memo = self._keyerr_as_emptystring(row, 'Point Memo')
            point.panel_name = self._keyerr_as_emptystring(row, 'Panel Name')
            point.out_of_service =  self._keyerr_as_emptystring(row, 'Out of Service')
            point.normal_ack_enabled =  self._keyerr_as_emptystring(row, 'Normal ack Enabled')
            point.night_mode_0 =  self._keyerr_as_emptystring(row, 'Night Mode (0)')
            point.mode_delay =  self._keyerr_as_emptystring(row, 'Mode Delay')
            point.low_alarm_limit =  self._keyerr_as_emptystring(row, 'Low Alarm Limit')
            point.level_delay =  self._keyerr_as_emptystring(row, 'Level Delay')
            point.intercept =  self._keyerr_as_emptystring(row, 'Intercept')
            point.initial_value =  self._keyerr_as_emptystring(row, 'Initial Value')
            point.initial_priority =  self._keyerr_as_emptystring(row, 'Initial Priority')
            point.informational_text =  self._keyerr_as_emptystring(row, 'Informational Text')
            point.high_alarm_limit =  self._keyerr_as_emptystring(row, 'High Alarm Limit')
            point.graphic_name =  self._keyerr_as_emptystring(row, 'Graphic Name')
            point.enhanced_alarms =  self._keyerr_as_emptystring(row, 'Enhanced Alarms')
            point.enhanced_alarm_mode_point =  self._keyerr_as_emptystring(row, 'Enhanced Alarm Mode Point')
            point.engineering_units =  self._keyerr_as_emptystring(row, 'Engineering Units')
            point.differential =  self._keyerr_as_emptystring(row, 'Differential')
            point.descriptor = self._keyerr_as_emptystring(row, 'Descriptor')
            point.day_mode_1 = self._keyerr_as_emptystring(row, 'Day Mode (1)')
            point.cov_limit = self._keyerr_as_emptystring(row, 'COV Limit')
            point.classification =  self._keyerr_as_emptystring(row, 'Classification')
            point.analog_representation =  self._keyerr_as_emptystring(row, 'Analog Representation')
            point.alarmable = self._keyerr_as_emptystring(row, 'Alarmable')
            point.alarm_message = self._keyerr_as_emptystring(row, 'Alarm Message')
            point.alarm_destinations = self._keyerr_as_emptystring(row, 'Alarm Destinations')
            point.aim = self._keyerr_as_emptystring(row, 'AIM')
            point.address_type = self._keyerr_as_emptystring(row, 'Address Type')
            point.actuator_type = self._keyerr_as_emptystring(row, 'Actuator Type')
            point.number_of_decimal_places = self._keyerr_as_emptystring(row, '# of decimal places')

    def analyze(self):
        for point in self.point_list:
            print point.name, point.system_name
            if point.name != point.system_name:
                try:
                    self.analysis["Name Mismatch"].append(point)
                except KeyError:
                    self.analysis["Name Mismatch"] = [point]

    def _keyerr_as_emptystring(self, dictionary, key):
        try:
            return dictionary[key]
        except KeyError:
            print dictionary, key, "Not Found"
            return ''


