import json, sys, subprocess, os, time


with open("domain_servers.json", "r") as fp:
    data = json.load(fp)
    fp.close()

def config_generator(site):
    config = """[Configuration]
Domain=goodix.com
DomainDNS={}
DomainController={}
DomainAdmin=ithelp
DomainAdminPwd=NTY3OEBjb20=
{}
PCNameBeginWith=PC-
NBNameBeginWith=LP-
""".format(data[site]["ip"], data[site]["controller"], data[site]["OU"])
    with open("Config.ini", "w") as fp:
            fp.write(config)
            fp.close()
    return None


def powershell_cmds():
    # 解开powershell 执行脚本限制
    p = subprocess.Popen(["powershell.exe", "Set-ExecutionPolicy Unrestricted"])
    p.communicate()

    # 执行加域脚本
    args = ["powershell", os.path.join(os.getcwd(), 'ComputerJoinDomain.ps1')]
    p = subprocess.Popen(args, stdout=subprocess.PIPE)
    # 实时打印加域脚本执行情况
    with open("log.txt", 'wb') as fp:
        for c in iter(p.stdout.readline, b''):
            fp.write(c)
        fp.close()
    p.stdout.close()
    p.wait()

def text_choicer():
    while True:
        op1 = int(input("""
        Please choose where are you working:
        1) Asia
        2) America
        3) Europe
        4) Egypt
        Please input your choice below:
    """))
        if op1 in (1, 2, 3, 4):
            if op1 == 1:
                op2 = int(input("""
        Please indicate which office you are working in:
        1) China
        2) Korea
        3) India        
    """))
                if op2 == 1:
                    op3 = int(input("""
        请选择办公室：
        1) 深圳
        2) 成都
        3) 上海
        4) 武汉
        5) 台湾
    """))
                    if op3 == 1:
                        print("*" * 30)
                        print("深圳办公室相关参数生成中...")
                        config_generator('SZ')
                        print("配置生成完成，请查看:")
                        with open ('Config.ini', 'r') as f:
                            for line in f:
                                print(line)
                        print("*" * 30)
                        time.sleep(5)
                        break
                    elif op3 == 2:
                        print("*" * 30)
                        print("成都办公室相关参数生成中...")
                        config_generator('CD')
                        print("配置生成完成，请查看:")
                        with open ('Config.ini', 'r') as f:
                            for line in f:
                                print(line)
                        print("*" * 30)
                        time.sleep(5)
                        break
                    elif op3 == 3:
                        print("*" * 30)
                        print("上海办公室相关参数生成中...")
                        config_generator('SH')
                        print("配置生成完成，请查看:")
                        with open ('Config.ini', 'r') as f:
                            for line in f:
                                print(line)
                        print("*" * 30)
                        time.sleep(5)
                        break
                    elif op3 == 4:
                        print("*" * 30)
                        print("武汉办公室相关参数生成中...")
                        config_generator('WH')
                        print("配置生成完成，请查看:")
                        with open ('Config.ini', 'r') as f:
                            for line in f:
                                print(line)
                        print("*" * 30)
                        time.sleep(5)
                        break
                    elif op3 == 5:
                        print("*" * 30)
                        print("台灣辦公室相關參數生成中...")
                        config_generator('TW')
                        print("配置生成完成，請查看:")
                        with open ('Config.ini', 'r') as f:
                            for line in f:
                                print(line)
                        print("*" * 30)
                        time.sleep(5)
                        break
                    else:
                        print("-"* 30)
                        print("Please choose correct options")
                        print("-"* 30)
                        time.sleep(5)
                        continue
                elif op2 in (2, 3):
                    op3 = int(input("""
        Are you VAS member?
        1) yes
        2) no
    """))
                    if op3 == 1:
                        print("*" * 30)
                        print("Related office config is generating...")
                        config_generator('KR-vas')
                        print("Config generated, please check:")
                        with open ('Config.ini', 'r') as f:
                            for line in f:
                                print(line)
                        print("*" * 30)
                        time.sleep(5)
                        break
                    elif op3 == 2:
                        print("*" * 30)
                        print("Related office config is generating...")
                        config_generator('KR')
                        print("Config generated, please check:")
                        with open ('Config.ini', 'r') as f:
                            for line in f:
                                print(line)
                        print("*" * 30)
                        time.sleep(5)
                        break
                    else:
                        print("-"* 30)
                        print("Please choose correct options")
                        print("-"* 30)
                        time.sleep(5)
                        continue
                else:
                    print("-"* 30)
                    print("Please choose correct options")
                    print("-"* 30)
                    time.sleep(5)
                    continue
            elif op1 == 2:
                op2 = int(input("""
        Please indicate which office you are working in:
        1) Austin
        2) Irvine        
    """))
                if op2 == 1:
                    
                    print("*" * 30)
                    print("Austin office config is generating...")
                    config_generator('US-Austin')
                    print("Config generated, please check:")
                    with open ('Config.ini', 'r') as f:
                        for line in f:
                            print(line)
                    print("*" * 30)
                    time.sleep(5)
                    break
                elif op2 == 2:
                    
                    print("Irvine office config is generating...")
                    config_generator('US-Irvine')
                    print("Config generated, please check:")
                    with open ('Config.ini', 'r') as f:
                        for line in f:
                            print(line)
                    print("*" * 30)
                    time.sleep(5)
                    break
                else:
                    print("-"* 30)
                    print("Please choose correct options")
                    print("-"* 30)
                    time.sleep(5)
                    continue
            elif op1 == 3:
                op2 = int(input("""
        Please choose below options:
        1) VAS menbers - Leuven, Nijmegen, Valbonne
        2) Commsolid members - Germany
    """))
                if op2 == 1:
                    print("*" * 30)
                    print("Related office config is generating...")
                    config_generator("NL")
                    print("Config generated, please check:")
                    with open ('Config.ini', 'r') as f:
                        for line in f:
                            print(line)
                    print("*" * 30)
                    time.sleep(5)
                    break
                elif op2 == 2:
                    print("*" * 30)
                    print("Related office config is generating...")
                    config_generator("GR-Commsolid")
                    print("Config generated, please check:")
                    with open ('Config.ini', 'r') as f:
                        for line in f:
                            print(line)
                    print("*" * 30)
                    time.sleep(5)
                    break
                else:
                    print("-"* 30)
                    print("Please choose correct options")
                    print("-"* 30)
                    time.sleep(5)
                    continue
            elif op1 == 4:
                
                print("*" * 30)
                print("Egypt office config is generating...")
                config_generator("Egypt")
                print("Config generated, please check:")
                with open ('Config.ini', 'r') as f:
                    for line in f:
                        print(line)
                print("*" * 30)
                time.sleep(5)
                break
        else:
            print("-"* 30)
            print("Please choose correct options")
            print("-"* 30)
            time.sleep(5)
            continue



if __name__ == "__main__":
    text_choicer()
    print("*" * 30)
    while True:
        op_1 = input("Is the configuration right for your environment? y/n\n")
        if op_1.lower() == 'y':
            print("running domain join script...")
            powershell_cmds()
            print("The script is executed, please check log.txt if there is any problem.")
            print("*" * 30)
            time.sleep(5)
            break
        else:
            print("Please re-generate config file...")
            text_choicer()
    exit