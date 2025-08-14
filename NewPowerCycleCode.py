# API details for power cycling
apiUrl = "http://172.20.97.2/rps"
apiUser = "root"
apiPass = "root"
baudRate = 115200
iteration = 50

# Configuration values (replace with your desired values)
rebootCount = 10  # Number of reboots

# Serial device
serialDevice = "/dev/ttyUSB0"
username = "ubuntu"
password = "ubuntu123"

# Interfaces & IPs (Hardcoded)
ipAddresses = ["192.168.1.11", "192.168.2.12", "192.168.3.13", "192.168.4.14", "192.168.5.15", "192.168.6.16", "192.18.7.17", "192.168.8.18"]
interfaces = ["eno1", "enp13s0", "enp14s0", "enp15s0", "enp16s0f0", "enp16s0f1", "enp16s0f2", "enp16s0f3"]

# Set up logging
logFilename = "consoleLog2.txt"
logging.basicConfig(filename=logFilename, level=logging.INFO, format="%(asctime)s - %(message)s")

with open(logFilename, "w") as logFile:
    logFile.write("Starting new log...\n")


def logMessage(message):
    """Prints and logs messages to consoleLog.txt"""
    print(message)
    with open(logFilename, "a") as logFile:
        logFile.write(message + "\n")


def cleanOutput(output):
    """Removes unwanted escape sequences from serial output"""
    output = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]', '', output)
    output = re.sub(r'\x1b\[\?2004[hl]', '', output)
    return output.strip()

def readSerialOutput(ser, timeout=15, command_sent=None, remove_extra_shell=True):
    output = ""
    start_time = time.time()
    while time.time() - start_time < timeout:
        if ser.inWaiting() > 0:
            output += ser.read(ser.inWaiting()).decode(errors="ignore")
        time.sleep(0.1)
    output = cleanOutput(output)

    if command_sent:
        #remove the command and any preceeding or proceeding whitespace.
        output = re.sub(rf"^\s*{re.escape(command_sent)}\s*$", "", output, flags=re.MULTILINE).strip()

    if remove_extra_shell:
        output = re.sub(r"^\s*Shell>\s*$", "", output, flags=re.MULTILINE).strip()
    return output

def extractSystemStats(output):
    #Extracts CPU Usage, Temperature, and Memory Usage from console output
    cpuUsage, temperature, memoryUsage = "Not Found", "Not Found", "Not Found"
    output = output.replace("\r", "").replace("\n", "").replace(" ", " ")
    match = re.search(r"Usage of /:\s+([\d.]+%)", output)
    if match:
        cpuUsage = match.group(1)
    match = re.search(r"Temperature:\s+([\d.]+ C)", output)
    if match:
        temperature = match.group(1)
    match = re.search(r"Memory usage:\s+([\d.]+%)", output)
    if match:
        memoryUsage = match.group(1)
    return cpuUsage, temperature, memoryUsage

def parseFreeOutput(output, prefix):
    """Parses the output of the free -h command."""
    lines = output.splitlines()
    if len(lines) < 2:
        return {}

    headers = lines[0].split()
    data = lines[1].split()

    result = {}
    for i, header in enumerate(headers):
        if i < len(data):
            result[f"{prefix} mem {header.lower()}"] = data[i]

    if len(lines) > 2:
        swapHeaders = lines[0].split()
        swapData = lines[2].split()
        for i, header in enumerate(swapHeaders):
            if i < len(swapData):
                result[f"{prefix} swap {header.lower()}"] = swapData[i]
    return result

# Ensure Excel file exists
excelFile = "testResults2.xlsx"
if os.path.exists(excelFile):
    os.remove(excelFile)  # Delete old file

wb = Workbook()
ws = wb.active

# Start with a minimal header (interfaces will be added later)
header = ["Iteration", "Timestamp", "CPU Usage", "CPU Temperature", "Memory Usage", "Power Cycle Status", "USB Status", "GPS Status"]
# Add interface columns to header
headerExtend = ["LTE1 Status", "LTE2 Status", "SIM1 Status", "SIM2 Status", "Modem1 Connection Status",
                "Modem2 Connection Status", "SIM1 Operator Name", "SIM2 Operator Name", "SIM1 Registration",
                "SIM2 Registration"]
for iface in interfaces:
    headerExtend.extend([f"{iface} Ethernet Detected", f"{iface} Speed", f"{iface} Ping Result"])

sensorHeaders = []
freeHeaders = []
mpstatHeaders = []

headerFull = header + headerExtend + sensorHeaders + freeHeaders + mpstatHeaders

ws.append(headerFull)

# DYNAMIC COLUMN WIDTH ADJUSTMENT
for colNum, headerItem in enumerate(headerFull, 1):
    columnLetter = get_column_letter(colNum)
    maxLength = len(headerItem)  # Start with header length

    # Check data in the column for longer values
    for row in ws.iter_rows(min_row=2):
        cellValue = row[colNum - 1].value
        if cellValue and len(str(cellValue)) > maxLength:
            maxLength = len(str(cellValue))

    # Add some padding
    ws.column_dimensions[columnLetter].width = maxLength + 5

# ADDITIONAL FORMATTING
# Header formatting
headerFont = Font(bold=True)
headerAlignment = Alignment(horizontal='center', vertical='center')

for cell in ws[1]:  # Apply to the header row
    cell.font = headerFont
    cell.alignment = headerAlignment

# Cell alignment for data rows
dataAlignment = Alignment(horizontal='center', vertical='center')

# Run for multiple iterations
for ite in range(1, iteration + 1):
    logData = []
    logMessage(f"\n========== Iteration {ite} ==========\n")
    logData.append(f"\n========== Iteration {ite} ==========\n")

    powerCycleStatus = "Success"  # Initialize as success, will change if any reboot fails
    '''
    # Run for multiple reboots
    for rebootNum in range(1, rebootCount + 1):
        logMessage(f"\n========== Reboot {rebootNum}/{rebootCount} ==========\n")

        # Determine on and off durations for this iteration
        onDuration = rebootNum

        # Turn OFF the port
        logMessage(f"Turning OFF the port (Reboot {rebootNum})...")
        try:
            offResponse = requests.get(apiUrl, params={"set_switch": "3 false"}, auth=HTTPBasicAuth(apiUser, apiPass), timeout=10)
            if offResponse.status_code == 200:
                logMessage("Port turned OFF successfully.")
            else:
                logMessage(f"Failed to turn OFF port. Status code: {offResponse.status_code}")
                powerCycleStatus = "Failed"  # Update status if any reboot fails
        except requests.exceptions.RequestException as e:
            logMessage(f"Failed to turn OFF port. Error: {e}")
            powerCycleStatus = "Failed"  # Update status if any reboot fails

        time.sleep(0.9)

        # Turn ON the port
        logMessage(f"Turning ON the port (Reboot {rebootNum})...")
        try:
            onResponse = requests.get(apiUrl, params={"set_switch": "3 true"}, auth=HTTPBasicAuth(apiUser, apiPass), timeout=10)
            if onResponse.status_code == 200:
                logMessage("Port turned ON successfully.")
            else:
                logMessage(f"Failed to turn ON port. Status code: {onResponse.status_code}")
                powerCycleStatus = "Failed"  # Update status if any reboot fails
        except requests.exceptions.RequestException as e:
            logMessage(f"Failed to turn ON port. Error: {e}")
            powerCycleStatus = "Failed"  # Update status if any reboot fails

        time.sleep(onDuration)
    '''
    logMessage("Reboot cycle completed.")
    
    logMessage("Waiting for 240 seconds for device reboot...")
    time.sleep(240)

    # Establish serial connection
    logMessage("Re-establishing serial connection...")
    try:
        ser = serial.Serial(serialDevice, baudRate, timeout=1)
        time.sleep(2)
    except serial.SerialException as e:
        logMessage(f"Error opening serial port: {e}")
        continue

    # Ensure login prompt
    logMessage("Checking for'ubuntu login:' prompt...")
    ser.write(b"\n" * 5)
    output = readSerialOutput(ser, 30, remove_extra_shell=False)
    if "ubuntu login:" not in output:
        logMessage("[ERROR] 'ubuntu login:' not detected, skipping this iteration.")
        continue

    # Enter credentials
    logMessage("Logging in...")
    ser.write(f"{username}\n".encode())
    output = readSerialOutput(ser, 10, remove_extra_shell=False)
    if "Password:" not in output:
        logMessage("[ERROR] Password prompt not detected, skipping this iteration.")
        continue
    ser.write(f"{password}\n".encode())
    output = readSerialOutput(ser, 10)

    # Check for login prompt
    if "root@ubuntu:" not in output and "ubuntu@ubuntu:" not in output:
        logMessage("[ERROR] Login failed, skipping this iteration.")
        continue

    # Extract system stats
    cpuUsage, temperature, memoryUsage = extractSystemStats(output)
    logMessage(f"CPU Usage: {cpuUsage}, CPU Temperature: {temperature}, Memory Usage: {memoryUsage}")

    ser.write("sudo su\n".encode())
    time.sleep(2)
    ser.write(f"{password}\n".encode())
    time.sleep(2)
    
    ser.write("sudo busybox devmem 0xFD6E0780 w 0x01\n".encode())
    time.sleep(2)
    
    # Clear the sensor data before each iteration.
    sensorsData = {}
    freeData = {}
    mpstatData = {}

    sensorHeaders = []
    freeHeaders = []
    mpstatHeaders = []

    # Sensor data collection (before stress-ng)
    logMessage("Collecting sensor data (before stress-ng)...")
    ser.write(b"sensors\n")
    sensorsOutput = readSerialOutput(ser, 10)
    logMessage(sensorsOutput)

    lines = sensorsOutput.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        if ":" in line and "Adapter:" not in line and "[sudo]" not in line and "root@ubuntu" not in line:
            parts = line.split(":")
            sensorName = parts[0].strip()
            sensorValue = parts[1].split("(")[0].strip()

            fullSensorName = sensorName
            if i > 0 and "Adapter:" in lines[i - 1]:
                fullSensorName = "before stress-ng " + lines[i - 2].strip() + " " + lines[i - 1].split(":")[1].strip() + " " + sensorName
            elif i > 0 and "Core" in sensorName:
                fullSensorName = "before stress-ng " + sensorName
            else:
                fullSensorName = "before stress-ng " + sensorName

            sensorsData[fullSensorName] = sensorValue
        i += 1

    # Memory usage details (free -h) (before stress-ng)
    logMessage("Collecting memory usage details (free -h) (before stress-ng)...")
    command = f"echo {password} | sudo -S free -h"
    ser.write(f"{command}\n".encode())
    freeOutput = readSerialOutput(ser, 10, command_sent=command)
    logMessage(freeOutput)

    freeData.update(parseFreeOutput(freeOutput, "before stress-ng"))

    # mpstat -P ALL 1 1 (before stress-ng)
    logMessage("Collecting mpstat -P ALL 1 1 data (before stress-ng)...")
    ser.write(b"mpstat -P ALL 1 1\n")
    mpstatOutput = readSerialOutput(ser, 10)
    logMessage(mpstatOutput)

    mpstatLines = mpstatOutput.splitlines()
    for line in mpstatLines:
        if line.startswith("Average:") or line.startswith("1"):
            parts = line.split()
            if len(parts) >= 12 and parts[1] != "CPU":  # Add this condition
                cpuName = parts[1]
                mpstatDetails = {
                    f"before stress-ng average CPU {cpuName} %usr": parts[2],
                    f"before stress-ng average CPU {cpuName} %nice": parts[3],
                    f"before stress-ng average CPU {cpuName} %sys": parts[4],
                    f"before stress-ng average CPU {cpuName} %iowait": parts[5],
                    f"before stress-ng average CPU {cpuName} %irq": parts[6],
                    f"before stress-ng average CPU {cpuName} %soft": parts[7],
                    f"before stress-ng average CPU {cpuName} %steal": parts[8],
                    f"before stress-ng average CPU {cpuName} %guest": parts[9],
                    f"before stress-ng average CPU {cpuName} %gnice": parts[10],
                    f"before stress-ng average CPU {cpuName} %idle": parts[11],
                }
                mpstatData.update(mpstatDetails)

    # stress-ng commands
    logMessage("Starting stress-ng commands...")
    ser.write(b"nohup stress-ng --vm $(nproc) --vm-bytes 100% --timeout 14m &\n")
    time.sleep(2)  # give time for command to start.
    ser.write(b"nohup stress-ng --cpu $(nproc) --timeout 14m &\n")
    time.sleep(5)  # give time for command to start.
    logMessage("stress-ng commands started.")

    # Sensor data collection (after stress-ng)
    logMessage("Collecting sensor data (after stress-ng)...")
    ser.write(b"sensors\n")
    sensorsOutput = readSerialOutput(ser, 10)
    logMessage(sensorsOutput)

    lines = sensorsOutput.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        if ":" in line and "Adapter:" not in line and "[sudo]" not in line and "root@ubuntu" not in line:
            parts = line.split(":")
            sensorName = parts[0].strip()
            sensorValue = parts[1].split("(")[0].strip()

            fullSensorName = sensorName
            if i > 0 and "Adapter:" in lines[i - 1]:
                fullSensorName = "after stress-ng " + lines[i - 2].strip() + " " + lines[i - 1].split(":")[1].strip() + " " + sensorName
            elif i > 0 and "Core" in sensorName:
                fullSensorName = "after stress-ng " + sensorName
            else:
                fullSensorName = "after stress-ng " + sensorName

            sensorsData[fullSensorName] = sensorValue
        i += 1

    # Memory usage details (free -h) (after stress-ng)
    logMessage("Collecting memory usage details (free -h) (after stress-ng)...")
    command = f"echo {password} | sudo -S free -h"
    ser.write(f"{command}\n".encode())
    freeOutput = readSerialOutput(ser, 10, command_sent=command)
    logMessage(freeOutput)

    freeData.update(parseFreeOutput(freeOutput, "after stress-ng"))

    # mpstat -P ALL 1 1 (after stress-ng)
    logMessage("Collecting mpstat -P ALL 1 1 data (after stress-ng)...")
    ser.write(b"mpstat -P ALL 1 1\n")
    mpstatOutput = readSerialOutput(ser, 10)
    logMessage(mpstatOutput)

    mpstatLines = mpstatOutput.splitlines()
    for line in mpstatLines:
        if line.startswith("Average:") or line.startswith("1"):
            parts = line.split()
            if len(parts) >= 12 and parts[1] != "CPU":  # Add this condition
                cpuName = parts[1]
                mpstatDetails = {
                    f"after stress-ng average CPU {cpuName} %usr": parts[2],
                    f"after stress-ng average CPU {cpuName} %nice": parts[3],
                    f"after stress-ng average CPU {cpuName} %sys": parts[4],
                    f"after stress-ng average CPU {cpuName} %iowait": parts[5],
                    f"after stress-ng average CPU {cpuName} %irq": parts[6],
                    f"after stress-ng average CPU {cpuName} %soft": parts[7],
                    f"after stress-ng average CPU {cpuName} %steal": parts[8],
                    f"after stress-ng average CPU {cpuName} %guest": parts[9],
                    f"after stress-ng average CPU {cpuName} %gnice": parts[10],
                    f"after stress-ng average CPU {cpuName} %idle": parts[11],
                }
                mpstatData.update(mpstatDetails)

    # Update sensor headers and data
    sensorHeaders = list(sensorsData.keys())
    freeHeaders = list(freeData.keys())
    mpstatHeaders = list(mpstatData.keys())

    headerFull = header + headerExtend + sensorHeaders + freeHeaders + mpstatHeaders

    #Update header in excel
    for colNum, headerItem in enumerate(headerFull, 1):
        ws.cell(row=1, column=colNum, value=headerItem)

    # eMMC command
    logMessage("Starting eMMC command...")
    ser.write(b"./eMMC_aggressive_3mins.sh &\n")
    time.sleep(2) # give time for command to start.

    # Configure network interfaces
    logMessage("Configuring Network Interfaces...")

    # Configure network interfaces with retry mechanism
    interfaceResults = {}
    for i, iface in enumerate(interfaces):
        logMessage(f"Attempting to bring up interface {iface}...")
        for retry in range(3):
            ser.write(f"echo {password} | sudo -S ip link set {iface} up\n".encode())
            time.sleep(5)  # Increased sleep time

            if i < len(ipAddresses):
                logMessage(f"Assigning IP {ipAddresses[i]} to {iface}...")
                ser.write(f"echo {password} | sudo -S ip addr add {ipAddresses[i]}/24 dev {iface}\n".encode())
                time.sleep(5)  # Increased sleep time

            logMessage(f"Checking operational state of {iface} (attempt {retry + 1})...")
            ser.write(f"echo {password} | sudo -S ethtool {iface}\n".encode())
            time.sleep(5)  # Increased sleep time

            ethtoolOutput = readSerialOutput(ser, 5)

            linkDetectedMatch = re.search(r"Link detected: yes", ethtoolOutput)
            speedMatch = re.search(r"Speed:\s*(\d+)Mb/s", ethtoolOutput)

            linkStatus = "Yes" if linkDetectedMatch else "No"
            speed = speedMatch.group(0) if speedMatch else "Unknown"

            logMessage(f"Interface {iface} - Link Detected: {linkStatus}, Speed: {speed}")

            if linkDetectedMatch:
                if retry > 0:
                    logMessage(f"Interface {iface} came up after {retry+1} retries in iteration {ite}")
                break  # Interface came up, break retry loop
            elif retry == 2:  # All retries failed
                logMessage(f"Interface {iface} failed to come up after 3 retries in iteration {ite}")
                break

        if linkDetectedMatch and i < len(ipAddresses):
            logMessage(f"Pinging {ipAddresses[i]} from {iface}...")
            ser.write(f"ping -c 10 {ipAddresses[i]}\n".encode())
            time.sleep(15)

            pingOutput = readSerialOutput(ser, 15)
            logMessage(pingOutput)

            match = re.search(r"(\d+) packets transmitted, (\d+) received, (\d+)% packet loss", pingOutput)

            if match:
                _, packetsReceived, packetLoss = map(int, match.groups())
                pingResult = "Ping Passed" if packetLoss == 0 else "Ping Failed"
            else:
                pingResult = "Ping Failed"

            logMessage(f"Ping Result for {iface}: {pingResult}")
        elif linkStatus == "No":
            pingResult = "Link Down"
        else:
            pingResult = "Speed Not 1000Mb/s"

        # Update interfaceResults for excel sheet.
        if iface not in interfaceResults:
            interfaceResults[iface] = {}
        interfaceResults[iface]["Ethernet Detected"] = linkStatus
        interfaceResults[iface]["Speed"] = speed
        interfaceResults[iface]["Ping Result"] = pingResult

    # Step 7: USB Check
    logMessage("Checking for USB device (/dev/sda)...")
    ser.write(b"ls /dev/sd*\n")
    usbOutput = readSerialOutput(ser, 5)
    logMessage(usbOutput)
    usbStatus = "USB Found" if "/dev/sda" in usbOutput else "USB Not Found"
    logMessage(f"USB Status: {usbStatus}")

    # Step 10: Check the interface was created for the modules (for both LTEs)
    lteInterfaces = ["LTE Interface Not Found", "LTE Interface Not Found"]  # prefill with default values.
    lteDevicesCount = 0  # count the Telit devices.
    for lteNum in range(2):  # Assuming 2 LTEs, adjust as needed
        # Step 8: LTE (Long Term Evolution) Check
        logMessage("Checking for LTE")
        ser.write(b"lsusb\n")
        lteOutput = readSerialOutput(ser, 5)
        logMessage(lteOutput)
        lteStatus = "LTE Found" if "Telit Wireless Solutions" in lteOutput else "LTE Not Found"
        if "Telit Wireless Solutions" in lteOutput:
            lteDevicesCount += 1
        logMessage(f"LTE Status: {lteStatus}")
        logMessage(f"Checking LTE interface {lteNum + 1} was created or not")
        ser.write(b"ip link show\n")
        interfaceOutput = readSerialOutput(ser, 5)
        logMessage(interfaceOutput)

        logMessage(f"LTE Interface {lteNum + 1}: {lteInterfaces[lteNum]}")  # display the default value.

    # Step 12: Checking modem status (for both modems)
    modemConnectionStatuses = []
    for modemNum in range(2):  # Assuming 2 modems, adjust as needed
        logMessage(f"Checking Modem {modemNum + 1} Status...")

        # Determine the correct modem index (always 0 or 1)
        modemIndex = str(modemNum)

        # Attempt to connect the modem with retry
        connectionStatus = "Connection Failed"
        for retry in range(3):
            ser.write(f"mmcli -m {modemIndex}\n".encode())
            modemOutput = readSerialOutput(ser, 5)

            connectCommand = f"sudo mmcli -m {modemIndex} --simple-connect=\"apn=airtelgprs.com\"\n"
            ser.write(connectCommand.encode())
            modemStatusOutput = readSerialOutput(ser, 5)
            if "password for ubuntu" in modemStatusOutput.lower():
                logMessage("Sudo password required, sending password...")
                ser.write(f"{password}\n".encode())  # Sending the password
                modemStatusOutput = readSerialOutput(ser, 5)  # Read the output again

            logMessage(modemStatusOutput)

            # Check for successful connection
            if "successfully connected the modem" in modemStatusOutput.lower():
                connectionStatus = "Connected Successfully"
                break
            time.sleep(3)

        logMessage(f"Modem {modemNum + 1} Connection Status: {connectionStatus}")
        modemConnectionStatuses.append(connectionStatus)

    # Step 11: Identify the USB Ports Mapped for AT Command Communication and Verify the Module Loaded for the SIM (for both SIMs)
    simStatuses = []
    for simNum in range(2):  # Assuming 2 SIMs, adjust as needed
        logMessage(f"Checking SIM {simNum + 1} Status...")

        # Send mmcli command (always -m 0 or -m 1)
        ser.write(f"mmcli -m {simNum}\n".encode())
        simOutput = readSerialOutput(ser, 5)

        logMessage(simOutput)

        # Default values
        simStatus = "SIM Not Detected"
        simSlot = "Unknown"

        # Extract SIM Slot & Status
        match = re.search(r"slot (\d+):\s*(.+?)\s*\(active\)", simOutput)
        if match:
            simSlot = f"Slot {match.group(1)}"
            if "none" in match.group(2).lower():
                simStatus = "SIM Not Detected"
            else:
                simStatus = "SIM Active"

        print(f"SIM Slot: {simSlot}, SIM Status: {simStatus}")
        simStatuses.append({"Slot": simSlot, "Status": simStatus})

    # Step 13: Get the sim details (for both SIMs)
    simDetails = []
    for modemNum in range(2):  # Assuming 2 modems, adjust as needed
        logMessage(f"Getting SIM {modemNum + 1} Details...")

        # Determine the correct modem index (always 0 or 1)
        modemIndex = str(modemNum)
        ser.write(f"mmcli -m {modemIndex}\n".encode())
        modemOutput = readSerialOutput(ser, 5)

        # Run command to get SIM details
        getSimCommand = f"sudo mmcli -m {modemIndex}\n"
        ser.write(getSimCommand.encode())

        # Check if sudo password is required
        simDetailsOutput = readSerialOutput(ser, 5)
        if "password for ubuntu" in simDetailsOutput.lower():
            logMessage("Sudo password required, sending password...")
            ser.write(f"{password}\n".encode())
            simDetailsOutput = readSerialOutput(ser, 5)

        logMessage(simDetailsOutput)

        # Extract the required details
        operatorName = f"SIM{modemNum + 1} Operator Name Not Found"
        registrationStatus = f"SIM{modemNum + 1} Registration Not Found"

        details = {"Operator Name": operatorName, "Registration": registrationStatus}
        if "operator name" in simDetailsOutput.lower() and "registration" in simDetailsOutput.lower():
            try:
                for line in simDetailsOutput.split("\n"):
                    if "operator name" in line:
                        operatorName = line.split(":")[-1].strip()
                        details["Operator Name"] = operatorName
                    elif "registration" in line:
                        registrationStatus = line.split(":")[-1].strip()
                        details["Registration"] = registrationStatus
            except Exception as e:
                logMessage(f"Error parsing SIM {modemNum + 1} details: {e}")
        else:
            logMessage(f"SIM {modemNum + 1} operator name or registration not found")

        # Log the extracted values
        logMessage(f"SIM {modemNum + 1} Operator Name: {details['Operator Name']}, Registration: {details['Registration']}")
        simDetails.append(details)

    # Step GPS Check
    logMessage("Checking GPS Status...")
    ser.write(b"lsusb\n")
    gpsOutput = readSerialOutput(ser, 5)
    gpsStatus = "GPS Found" if "U-Blox AG u-blox GNSS receiver" in gpsOutput else "GPS Not Found"
    logMessage(f"GPS Status: {gpsStatus}")

    # Step 14: Store data in Excel
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rowData = [ite, timestamp, cpuUsage, temperature, memoryUsage, powerCycleStatus, usbStatus, gpsStatus]

    # Add LTE and SIM results to rowData
    if lteDevicesCount == 1 and len(lteInterfaces) == 1:
        lte1Status = "LTE Found"
        lte2Status = "LTE Not Found"
    elif lteDevicesCount == 2 and len(lteInterfaces) == 2:
        lte1Status = "LTE Found"
        lte2Status = "LTE Found"
    else:
        lte1Status = "LTE Not Found"
        lte2Status = "LTE Not Found"

    rowData.extend([lte1Status, lte2Status])
    rowData.extend([simStatuses[0]['Status'] if len(simStatuses) > 0 else "SIM1 Not Detected",
                    simStatuses[1]['Status'] if len(simStatuses) > 1 else "SIM2 Not Detected"])
    rowData.extend([modemConnectionStatuses[0] if len(modemConnectionStatuses) > 0 else "Modem1 Connection Failed",
                    modemConnectionStatuses[1] if len(modemConnectionStatuses) > 1 else "Modem2 Connection Failed"])
    rowData.extend([simDetails[0]['Operator Name'] if len(simDetails) > 0 else "SIM1 Operator Name Not Found",
                    simDetails[1]['Operator Name'] if len(simDetails) > 1 else "SIM2 Operator Name Not Found"])
    rowData.extend([simDetails[0]['Registration'] if len(simDetails) > 0 else "SIM1 Registration Not Found",
                    simDetails[1]['Registration'] if len(simDetails) > 1 else "SIM2 Registration Not Found"])

    # Add interface results to excel
    for iface in interfaces:
        rowData.extend([interfaceResults[iface]["Ethernet Detected"], interfaceResults[iface]["Speed"],
                        interfaceResults[iface]["Ping Result"]])

    # Add sensor data
    for sensor in sensorHeaders:
        rowData.append(sensorsData.get(sensor, "Sensor Value Not Found"))

    # Add free data
    for freeKey in freeHeaders:
        rowData.append(freeData.get(freeKey, "Sensor Value Not Found"))

    # Add mpstat data
    for mpstatKey in mpstatHeaders:
        rowData.append(mpstatData.get(mpstatKey, "Sensor Value Not Found"))

    ws.append(rowData)
    wb.save(excelFile)

    # Apply cell alignment for data rows
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = dataAlignment

    logMessage(f"Data for iteration {ite} written to Excel.")

# Close the serial port
if 'ser' in locals() and ser.is_open:
    ser.close()

logMessage("Script completed.")
