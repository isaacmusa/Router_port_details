import re
import pandas as pd

def extract_interface_details(text):
    # Define regular expressions for extracting information
    interface_pattern = re.compile(r'(\S+) current state\s+:\s+(\S+)\s*\nLine protocol current state\s+:\s+(\S+)')
    mtu_pattern = re.compile(r'The Maximum Transmit Unit is (\d+)')
    rx_power_pattern = re.compile(r'Rx Power: (.+?)dBm, Warning range: \[(.+?), (.+?)\]dBm')
    tx_power_pattern = re.compile(r'Tx Power: (.+?)dBm, Warning range: \[(.+?), (.+?)\]dBm')
    input_bandwidth_pattern = re.compile(r'Last 300 seconds input rate: (\d+) bits/sec')
    output_bandwidth_pattern = re.compile(r'Last 300 seconds output rate: (\d+) bits/sec')

    # Extract information using regular expressions
    interface_matches = interface_pattern.finditer(text)

    interface_data = []
    for match in interface_matches:
        interface_name = match.group(1)
        port_status = f"{match.group(2)}/{match.group(3)}"
        
        mtu_match = mtu_pattern.search(text, match.end())
        mtu = int(mtu_match.group(1))

        rx_power_match = rx_power_pattern.search(text, mtu_match.end())
        rx_power = f"{rx_power_match.group(1)} [{rx_power_match.group(2)}, {rx_power_match.group(3)}]" if rx_power_match else "N/A"

        tx_power_match = tx_power_pattern.search(text, rx_power_match.end()) if rx_power_match else None
        tx_power = f"{tx_power_match.group(1)} [{tx_power_match.group(2)}, {tx_power_match.group(3)}]" if tx_power_match else "N/A"

        input_bandwidth_match = input_bandwidth_pattern.search(text, tx_power_match.end()) if tx_power_match else None
        input_bandwidth = f"{input_bandwidth_match.group(1)} bits/sec" if input_bandwidth_match else "N/A"

        output_bandwidth_match = output_bandwidth_pattern.search(text, input_bandwidth_match.end()) if input_bandwidth_match else None
        output_bandwidth = f"{output_bandwidth_match.group(1)} bits/sec" if output_bandwidth_match else "N/A"

        interface_data.append({
            'Interface': interface_name,
            'Port Status': port_status,
            'MTU': mtu,
            'RX Power (Warning Range)': rx_power,
            'TX Power (Warning Range)': tx_power,
            'Input Bandwidth Utilization': input_bandwidth,
            'Output Bandwidth Utilization': output_bandwidth
        })

    return interface_data

# Read text from the file
with open(r'C:\Users\ISAAC AYEGBA\Documents\Desktop document\1\VBS Script\Kano002-NE40E-AR-A.txt', 'r') as file:
    text_data = file.read()

# Extract interface details
interfaces_data = extract_interface_details(text_data)

# Create a DataFrame
df = pd.DataFrame(interfaces_data)

# Write the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)
