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
        self.name = ''
        self.value = ''
        self.units = ''
        self.status = ''
        self.priority = ''

    def __repr__(self):
        return str([self.address, self.device, self.name, self.value, self.units, self.status, self.priority])


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
    def __init__(self, report_path):
         super(PanelPointLogReport, self).__init__(report_path)


PanelPPCLReport()