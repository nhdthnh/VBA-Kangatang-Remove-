from scapy.all import ARP, Ether, srp
import socket
def scan_wifi_network(ip_range):
    # Create an ARP request packet
    arp_request = Ether(dst="ff:ff:ff:ff:ff:ff") / ARP(pdst=ip_range)

    # Send the packet and capture the responses
    result = srp(arp_request, timeout=3, verbose=0)[0]

    # Extract the IP addresses from the responses
    ip_addresses = []
    for sent, received in result:
        ip_addresses.append(received.psrc)

    return ip_addresses

# Specify the IP range to scan
start_ip = '192.168.1.0'
end_ip = '192.168.254.254'

# Scan the Wi-Fi network within the specified IP range
used_ips = []
ip_parts = start_ip.split('.')
current_ip = ip_parts.copy()

# Loop through IP range
while current_ip != end_ip.split('.'):
    ip_range = '.'.join(current_ip) + '/32'
    used_ips.extend(scan_wifi_network(ip_range))

    # Increment IP address
    for i in range(3, -1, -1):
        current_ip[i] = str(int(current_ip[i]) + 1)
        if int(current_ip[i]) < 256:
            break
        else:
            current_ip[i] = '0'

def get_device_info(ip_address):
    # Create an ARP request packet
    arp_request = Ether(dst="ff:ff:ff:ff:ff:ff") / ARP(pdst=ip_address)

    # Send the packet and capture the response
    result = srp(arp_request, timeout=3, verbose=0)[0]

    # Extract the MAC address from the response
    if result:
        mac_address = result[0][1].hwsrc

        # Get the hostname using the IP address
        try:
            hostname = socket.gethostbyaddr(ip_address)[0]
        except socket.herror:
            hostname = "Unknown"

        return mac_address, hostname

    return None, None

# Print the used IP addresses
for ip in used_ips:
    print(ip)
    print(get_device_info(ip))