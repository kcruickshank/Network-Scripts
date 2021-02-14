
ip_one = input('Enter 1st WLC IP Address: ')
ip_two = input('Enter 2nd WLC IP Address: ')

if ip_two == '':
	ips = [ip_one]
else:
	ips = [ip_one,ip_two]

ip_hex = ''
type = 'f1'
wlcs='{0:x}'.format(4*len(ips)).zfill(2)

for ip in ips:
	for oct in ip.split('.'):
		ip_hex +='{0:x}'.format(int(oct)).zfill(2)
		

print('\noption 43: %s%s%s' % (type,wlcs,ip_hex))

