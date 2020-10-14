import json, sys, subprocess, os


with open("domain_servers.json", "r") as fp:
    data = json.load(fp)
    fp.close()


def switch_site(n):
    switcher = {
        1: "SZ",
        2: "US",
        3: "IN",
        4: "KR",
        5: "KR-T1"
    }
    return switcher.get(n, "Invalid Input!")


def config_generator(site):
# change your DomainAdmin and DomainAdminPwd with your profile
    config = """[Configuration]
Domain=goodix.com
DomainDNS={}
DomainController={}
DomainAdmin=....
DomainAdminPwd=....
{}
PCNameBeginWith=PC-
NBNameBeginWith=LP-
""".format(data[site]["ip"], data[site]["controller"], data[site]["OU"])
    with open("Config.ini", "w") as fp:
            fp.write(config)
            fp.close()
    return None


def powershell_cmds():
    # powershell script restrict
    p = subprocess.Popen(["powershell.exe", "Set-ExecutionPolicy Unrestricted"])
    p.communicate()

    # exec domain join powershell script
    args = ["powershell", os.path.join(os.getcwd(), 'ComputerJoinDomain.ps1')]
    p = subprocess.Popen(args, stdout=subprocess.PIPE)
    # log exec process
    with open("log.txt", 'wb') as fp:
        for c in iter(p.stdout.readline, b''):
            fp.write(c)
        fp.close()
    p.stdout.close()
    p.wait()
    p = subprocess.Popen(["powershell.exe", "Set-ExecutionPolicy Restricted"])
    p.communicate()


while True:
    op1 = int(input("""
Please choose where are you working:
1) Shenzhen&SH, China
2) Irvine, USA
3) Bangalore, India
4) Seoul, South Korea
Please input your choice below:
"""))

    if 1 <= op1 <= 4:
        if op1 == 4:
            op2 = int(input("""
Are you in T1 team or not?
1) Yes
2) No
Please input your choice below:
"""))       
            if op2 == 1:
                site = switch_site(5)
            elif op2 == 2:
                site = switch_site(4)
            else:
                print("Please input correct number!")
        else:        
            site = switch_site(op1)
        config_generator(site)
        powershell_cmds()
    else:
        os.system("cls")
        print("Please input correct number!")
