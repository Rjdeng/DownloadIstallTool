import os
import requests
import pandas as pd
import subprocess
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor
from rich.text import Text
from rich.console import Console

console = Console()

def check_device_online():
    try:
        print("检测设备中...")

        result = subprocess.run(["adb", "devices"], stdout=subprocess.PIPE)
        devices = result.stdout.decode().splitlines()
        devices = [line for line in devices if line.endswith("device")]

        if len(devices) == 0:
            console.print(Text(f"未检测到设备，请检查设备连接状态后重试", style="bold red"))
            return None
        elif len(devices) > 1:
            console.print(Text(f"检测到 {len(devices)} 台设备，请确保只连接一台设备后重试", style="bold red"))
            return None
        else:
            device_id = devices[0].split()[0]
            console.print(Text(f"检测到一台设备：{device_id}", style="bold green"))
            return device_id
    except Exception as e:
        console.print(Text(f"检测设备时出错: {e}", style="bold red"))
        return None

def download_app(app_name, url, download_dir, show_progress=True):
    try:
        print(f"正在下载 {app_name}...")
        file_path = os.path.join(download_dir, f"{app_name}.apk")
        
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()

        total_size = int(response.headers.get('content-length', 0))

        with open(file_path, 'wb') as apk_file:
            if show_progress:
                with tqdm(
                    desc=app_name,
                    total=total_size,
                    unit='B',
                    unit_scale=True,
                    unit_divisor=1024,
                    bar_format='\033[94m{l_bar}{bar}\033[0m {n_fmt}/{total_fmt} {unit} {rate_fmt}{postfix}'
                ) as bar:
                    for data in response.iter_content(chunk_size=1024):
                        apk_file.write(data)
                        bar.update(len(data))
            else:
                for data in response.iter_content(chunk_size=1024):
                    apk_file.write(data)

        print(f"{app_name} 下载完成, 保存在 {file_path}")
        return file_path, "下载成功"
    except Exception as e:
        print(f"下载 {app_name} 失败: {e}")
        return None, f"下载失败: {e}"

def install_app(device_id, apk_path, show_progress=True):
    try:
        print(f"正在安装 {apk_path}...")
        subprocess.run(["adb", "-s", device_id, "install", apk_path], check=True)
        print(f"{apk_path} 安装完成")
        return "安装成功"
    except subprocess.CalledProcessError as e:
        print(f"安装 {apk_path} 失败: {e}")
        return f"安装失败: {e}"
    
def update_excel_status(excel_file, app_name, url, download_status, install_status):
    # 读取现有的 Excel 文件
    df = pd.read_excel(excel_file)

    if app_name in df['应用名'].values:
        # 更新现有行
        index = df[df['应用名'] == app_name].index[0]
        df.at[index, '下载状态'] = download_status
        df.at[index, '安装状态'] = install_status
    else:
        # 创建新行并确保它是 DataFrame
        new_row = pd.DataFrame({
            '应用名': [app_name], 
            '下载链接': [url], 
            '下载状态': [download_status], 
            '安装状态': [install_status]
        })
        # 将新行添加到 DataFrame 中
        df = pd.concat([df, new_row], ignore_index=True)

    # 保存更新后的文件
    df.to_excel(excel_file, index=False)
    print(f"已更新 {app_name} 的状态为: {download_status}, {install_status}")

def download_and_install_apps(device_id, download_dir, parallel=True):
    excel_file = os.path.join(download_dir, "download.xlsx")
    if not os.path.exists(excel_file):
        print("未找到 download.xlsx 文件")
        return

    df = pd.read_excel(excel_file)
    
    total_downloads = len(df)
    download_success = 0
    download_failure = 0

    total_installs = 0
    install_success = 0
    install_failure = 0

    def process_app(index, row):
        nonlocal download_success, download_failure, install_success, install_failure, total_installs

        app_name = row['应用名']
        url = row['下载链接']
        
        # 下载应用
        apk_path, download_status = download_app(app_name, url, download_dir, show_progress=True)
        
        # 统计下载状态
        if "成功" in download_status:
            download_success += 1
        else:
            download_failure += 1
        
        install_status = "未安装"
        # 如果下载成功，继续安装应用
        if apk_path:
            install_status = install_app(device_id, apk_path, show_progress=True)
            total_installs += 1
            
            # 统计安装状态
            if "成功" in install_status:
                install_success += 1
            else:
                install_failure += 1

        update_excel_status(excel_file, app_name, url, download_status, install_status)

    if parallel:
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(process_app, index, row) for index, row in df.iterrows()]
            for future in futures:
                future.result()
    else:
        for index, row in df.iterrows():
            process_app(index, row)

    # 打印下载和安装结果
    console.print(Text(f"下载完成: 成功 {download_success} 个, 失败 {download_failure} 个, 总计 {total_downloads} 个", style="bold bright_cyan"))
    console.print(Text(f"安装完成: 成功 {install_success} 个, 失败 {install_failure} 个, 总计 {total_installs} 个", style="bold bright_cyan"))

def install_local_apks(device_id, apk_dir, excel_file, parallel=True):
    # 检查 Excel 文件是否存在，如果不存在，则创建一个
    if not os.path.exists(excel_file):
        print(f"{excel_file} 文件未找到，正在创建...")
        
        # 创建一个初始的 DataFrame
        df = pd.DataFrame(columns=['应用名', '下载链接', '下载状态', '安装状态'])
        
        # 保存到 Excel 文件
        df.to_excel(excel_file, index=False)
        print(f"{excel_file} 文件已创建。")
    
    apk_files = [f for f in os.listdir(apk_dir) if f.endswith(".apk")]

    if not apk_files:
        print("当前目录未找到任何 APK 文件")
        return

    df = pd.read_excel(excel_file)
    
    total_installs = len(apk_files)
    install_success = 0
    install_failure = 0

    def install_apk(apk_file):
        nonlocal install_success, install_failure
        apk_path = os.path.join(apk_dir, apk_file)
        install_status = install_app(device_id, apk_path, show_progress=True)

        app_name = os.path.splitext(apk_file)[0]
        url = df[df['应用名'] == app_name]['下载链接'].values[0] if app_name in df['应用名'].values else "本地安装"

        update_excel_status(excel_file, app_name, url, "已下载", install_status)
        
        if "成功" in install_status:
            install_success += 1
        else:
            install_failure += 1

        print(f"APK 文件 {apk_file} 的安装状态为: {install_status}")

    if parallel:
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(install_apk, apk_file) for apk_file in apk_files]
            for future in futures:
                future.result()
    else:
        for apk_file in apk_files:
            install_apk(apk_file)

    # 打印安装结果
    console.print(Text(f"安装完成: 成功 {install_success} 个, 失败 {install_failure} 个, 总计 {total_installs} 个", style="bold bright_cyan"))
    
def download_apps(device_id, download_dir, parallel=True):
    excel_file = os.path.join(download_dir, "download.xlsx")
    if not os.path.exists(excel_file):
        print("未找到 download.xlsx 文件")
        return

    df = pd.read_excel(excel_file)
    
    total_downloads = len(df)
    download_success = 0
    download_failure = 0

    def process_app(index, row):
        nonlocal download_success, download_failure

        app_name = row['应用名']
        url = row['下载链接']
        
        # 下载应用
        apk_path, download_status = download_app(app_name, url, download_dir, show_progress=True)
        
        # 统计下载状态
        if "成功" in download_status:
            download_success += 1
        else:
            download_failure += 1

        # 更新 Excel 文件，安装状态设置为 "未安装"
        update_excel_status(excel_file, app_name, url, download_status, "未安装")

    if parallel:
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(process_app, index, row) for index, row in df.iterrows()]
            for future in futures:
                future.result()
    else:
        for index, row in df.iterrows():
            process_app(index, row)

    # 打印下载结果
    console.print(Text(f"下载完成: 成功 {download_success} 个, 失败 {download_failure} 个, 总计 {total_downloads} 个", style="bold bright_cyan"))

def delete_all_apks(current_dir):
    """
    删除当前目录下的所有 APK 文件。
    """
    apk_files = [f for f in os.listdir(current_dir) if f.endswith('.apk')]
    
    if not apk_files:
        console.print("当前目录没有 APK 文件。", style="bright_yellow")
        return

    console.print(f"找到以下 APK 文件将被删除：{', '.join(apk_files)}", style="bright_red")
    confirmation = input("确认删除这些文件吗？输入 'y' 确认删除，或输入 'n' 取消删除：").strip().lower()

    if confirmation == 'y':
        for apk in apk_files:
            try:
                os.remove(os.path.join(current_dir, apk))
                console.print(f"已删除文件：{apk}", style="bright_green")
            except Exception as e:
                console.print(f"删除文件 {apk} 时出错：{e}", style="bright_red")
    elif confirmation == 'n':
        console.print("文件删除已取消。", style="bright_yellow")
    else:
        console.print("无效的输入，请重新输入。", style="bright_red")

def main():
    console = Console()
    
    console.print(f"有需求请联系(wx)：BBKRjdeng", style="bright_green")

    while True:
        console.print(Text("请选择操作：\n1. 下载应用\n2. 安装应用\n3. 下载并安装应用\n4. 删除本目录所有APK\nq. 退出", style="bright_yellow"))
        choice = input("输入选项编号：").strip().lower()

        if choice == 'q':
            console.print("程序已退出。", style="bright_red")
            break

        current_dir = os.getcwd()
        excel_file = os.path.join(current_dir, "download.xlsx")

        if choice == '1':
            # 只下载应用
            download_apps(None, current_dir, parallel=True)
        elif choice == '2':
            # 只安装应用
            device_id = check_device_online()
            if device_id:
                install_local_apks(device_id, current_dir, excel_file, parallel=True)
        elif choice == '3':
            # 下载并安装应用
            device_id = check_device_online()
            if device_id:
                download_and_install_apps(device_id, current_dir, parallel=True)
        elif choice == '4':
            # 删除本目录所有APK
            delete_all_apks(current_dir)
        else:
            console.print("无效选项，请重新选择。", style="bright_red")

if __name__ == "__main__":
    main()




