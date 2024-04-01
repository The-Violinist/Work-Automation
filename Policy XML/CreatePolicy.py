from os import system
policy_name = input("Please enter the name of the new policy: ")

port_list = []

def new_port():
    while True:
        is_range = input("1) Single Port\n2) Port Range\n >")
        if is_range == "1" or is_range == "2":
            break
        else: print("Enter a valid choice")
    system('cls')

    while True:
        port_protocol = input("1) TCP\n2) UDP\n>")
        if port_protocol == "1":
            port_protocol = "6"
            break
        elif port_protocol == "2":
            port_protocol = "17"
            break
        else: print("Enter a valid choice")

    system('cls')
    if is_range == "1":
        port_num = input("Port Number\n>")
        return [is_range, port_protocol, port_num]
    elif is_range == "2":
        port_num = input("Starting port\n>")
        system('cls')
        end_port = input("Ending port\n>")
        return [is_range, port_protocol, port_num, end_port]


while True:
    continue_loop = input("Add a port? 'y' to continue, or simply press enter to exit\n>")
    if continue_loop == "y":
        system('cls')
        query_result = new_port()
        port_list.append(query_result)
    elif continue_loop == "":
        break
    else:
        print("Enter a valid selection.")

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
            <name>{policy_name}</name>
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
    <service>{policy_name}</service>
    <builtin-service>false</builtin-service>
    <proxy-service>false</proxy-service>
    </service-additional-info>
    </service-additional-info-list>
    </ui-pm>
    </policy-view>
</profile>
<!-- DOCTYPE rs-profile SYSTEM "profile.dtd" -->
'''

xml_write = open(f"{policy_name}.xml", "w")
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

