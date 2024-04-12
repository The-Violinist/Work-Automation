import json 

port_list = []

with open('port_list') as list_file: 
    list = json.load(list_file) 

policy = list[0]
tcp_ports = list[1]
udp_ports = list[2]
tcp_range = list[3]
udp_range = list[4]

for item in tcp_ports:
    port_list.append(["1","6",str(item)])

for item in udp_ports:
    port_list.append(["1","17",str(item)])

if tcp_range != []:
    for item in tcp_range:
        port_list.append(["2","6",str(item[0]),str(item[1])])

if udp_range != []:
    for item in udp_range:
        port_list.append(["2","17",str(item[0]),str(item[1])])


policy_beginning = f'''<?xml version="1.0" encoding="UTF-8"?>
<profile>
    <product-grade/>
    <rs-version/>
    <using-cpm-profile>0</using-cpm-profile>
    <for-version>11.7</for-version>
    <xml-purpose>2</xml-purpose>
    <base-model/>
    <service-list>
        <service>
            <name>{policy}</name>
            <description/>
            <property>0</property>
            <proxy-type/>
            <service-item>'''

policy_end = f'''
            </service-item>
            <idle-timeout>0</idle-timeout>
        </service>
    </service-list>
    <policy-view>
    <ui-pm>
    <service-additional-info-list>
    <service-additional-info>
    <service>{policy}</service>
    <builtin-service>false</builtin-service>
    <proxy-service>false</proxy-service>
    </service-additional-info>
    </service-additional-info-list>
    </ui-pm>
    </policy-view>
</profile>
<!-- DOCTYPE rs-profile SYSTEM "profile.dtd" -->
'''

xml_write = open(f"{policy}.xml", "w")
xml_write.write(policy_beginning)

for port in port_list:
    if port[0] == "1":
        port_memeber = f'''
                <member>
                    <type>{port[0]}</type>
                    <protocol>{port[1]}</protocol>
                    <server-port>{port[2]}</server-port>
                </member>'''
    else:
        port_memeber = f'''
                <member>
                    <type>{port[0]}</type>
                    <protocol>{port[1]}</protocol>
                    <start-server-port>{port[2]}</start-server-port>
                    <end-server-port>{port[3]}</end-server-port>
                </member>'''
    xml_write.write(port_memeber)


xml_write.write(policy_end)
xml_write.close()

