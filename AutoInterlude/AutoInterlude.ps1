# Interludeのウィンドウを前面に表示するためのWin32 API呼び出し
Add-Type @"
    using System;
    using System.Linq;
    using System.Diagnostics;
    using System.Runtime.InteropServices;

    public class User32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsIconic(IntPtr hWnd);
    }

    public class InterludeWindowManager {
        public static void BringToFront(string processName) {
            var processes = Process.GetProcessesByName(processName);
            var process = processes.Cast<Process>().FirstOrDefault();

            if (process != null) {
                IntPtr mainWindowHandle = process.MainWindowHandle;

                if (User32.IsIconic(mainWindowHandle)) {
                    User32.ShowWindow(mainWindowHandle, 9); // SW_RESTORE
                }

                User32.SetForegroundWindow(mainWindowHandle);
            } else {
                Console.WriteLine(String.Format("{0} is not running.", processName));
            }
        }
    }
"@

# Interludeを前面に表示
[InterludeWindowManager]::BringToFront("Interlude")

# System.Windows.Formsを利用するための準備
Add-Type -AssemblyName System.Windows.Forms


Start-Sleep -Seconds 1
# 画面切り替え（Altキーを模倣）
[System.Windows.Forms.SendKeys]::SendWait("%")
Start-Sleep -Seconds 0.5

# 帳票設定画面表示（'c' と 'p' キーを模倣）
[System.Windows.Forms.SendKeys]::SendWait("c")
[System.Windows.Forms.SendKeys]::SendWait("p")

# エクスポートに向けてのTabとDownキー操作
[System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{DOWN}")
# Tabキーを複数回送信して、特定のフィールドにフォーカスを移動
1..6 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }
Start-Sleep -Seconds 0.5

# Enterキーを模倣して、選択肢を確定
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 0.5

# エクスポートさせる
[System.Windows.Forms.SendKeys]::SendWait("{TAB}{ENTER}")


# 印刷タブへの移動（TabとRightキーを模倣）
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
[System.Windows.Forms.SendKeys]::SendWait("{RIGHT}{RIGHT}{RIGHT}{RIGHT}")

Start-Sleep -Seconds 1

# ファイルパスとリネーム先のディレクトリパスを設定
$filePath = "\\ascsv10\開発\100_消費財卸\03074_武本ﾎｰﾑｽﾞ\ASPAC\PROGRAM\INTERD.TXT.INI"
$destinationDirectory = "\\ascsv10\開発\100_消費財卸\03074_武本ﾎｰﾑｽﾞ\ASPAC\MG10ENV\Interlude\test"


# ファイルから"PrinterName"を含む行を探し、"PrinterName="を削除してプリンター名を抽出
$buf = Get-Content $filePath | Where-Object { $_ -match "PrinterName" } | ForEach-Object { $_ -replace "PrinterName=", "" }

# 最終的なファイル名を設定（$bufの値に基づき、.ini拡張子を付加）
$finalFileName = "$buf.ini"

# 新しいファイルパスを生成（リネーム先ディレクトリと最終的なファイル名を結合）
$newFilePath = Join-Path -Path $destinationDirectory -ChildPath $finalFileName

# ファイル名を変更し、指定したディレクトリに移動
Move-Item -Path $filePath -Destination $newFilePath


