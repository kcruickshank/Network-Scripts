import netmiko
from netmiko.ssh_exception import AuthenticationException, SSHException, NetMikoTimeoutException
from netmiko import ConnectHandler


ip_address = "10.10.252.1"
username = "local_user"
password = "M1cr0Lab2003"

# Define a switch type
switch = {
    "device_type": "cisco_ios",
                "ip": ip_address,
                "username": username,
                "password": password,
}

net_connect = ConnectHandler(**switch)

net_connect.enable()

switchname = net_connect.send_command ("sh ver | i uptime")

print(switchname)

