@echo off
REM ==============================================
REM ウィンドウ表示設定_変数定義
setlocal ENABLEDELAYEDEXPANSION
color 0A
set KEY=HKEY_CURRENT_USER\SOFTWARE\VRChat\VRChat




REM ==============================================
REM 事前処理 レジストリBKUP

REM vbsYorN表示
echo BKUP 作成確認
echo Dim answer:answer = MsgBox("レジストリの事前Bkupを作成しますか？" ^& Chr(13) ^& "この処理にはフレンド人数に応じて数秒～数十秒かかる場合があります",vbYESNO,"実行確認"):WScript.Quit(answer) > %TEMP%\msgbox.vbs & %TEMP%\msgbox.vbs
REM YN結果_変数格納
set YN=%ERRORLEVEL%
REM MsgBox表示用_一時vbs_削除
timeout /t 0 >nul
del /Q %TEMP%\msgbox.vbs

REM IF分岐(Yes選択時にBKUP作成)
if %YN% equ 6 (
    REM 事前Bkup(既存ファイル上書き)
    echo BKUP 作成中・・・
    reg export "HKEY_CURRENT_USER\SOFTWARE\VRChat\VRChat" "%TEMP%\HKCU_S_V_V_bkup.reg" /y
    REM 完了表示
    echo BKUP 作成完了
    echo msgbox "レジストリの事前Bkupを作成しました",vbInformation,"BKUP作成完了" > %TEMP%\msgbox.vbs & %TEMP%\msgbox.vbs
    timeout /t 0 >nul
    del /Q %TEMP%\msgbox.vbs
)




REM ==============================================
REM メイン処理 レジストリ削除

REM vbsYorN表示
echo レジストリ削除 実行確認
echo Dim answer:answer = MsgBox("レジストリを修正します よろしいですか？",vbOKCancel,"実行確認"):WScript.Quit(answer) > %TEMP%\msgbox.vbs & %TEMP%\msgbox.vbs
REM YN結果_変数格納
set YN=%ERRORLEVEL%
REM MsgBox表示用_一時vbs_削除
timeout /t 0 >nul
del /Q %TEMP%\msgbox.vbs

REM IF分岐(Yes選択時にregキー削除)
if %YN% equ 1 (
    echo レジストリ削除 実行開始

    REM レジストリ削除(Code by Narazaka)
    echo レジストリ削除 削除対象の検索中・・・
    reg query %KEY% /f "FriendsPerLocation*" > VRChat_FriendsPerLocation_raw.txt
    find "FriendsPerLocation" < VRChat_FriendsPerLocation_raw.txt > VRChat_FriendsPerLocation_filtered.txt

    echo レジストリ削除 検索用一時ファイル 削除①
    del VRChat_FriendsPerLocation_raw.txt

    echo レジストリ削除 レジストリキー削除処理中・・・
    set /a counter=0
    for /f %%t in (VRChat_FriendsPerLocation_filtered.txt) do (

    set /a counter=counter+1
    echo レジストリを削除中・・・ !counter!件
        reg delete %KEY% /v "%%t" /f >nul
    )

    echo レジストリ削除 検索用一時ファイル 削除②
    del VRChat_FriendsPerLocation_filtered.txt

    REM 完了表示
    echo レジストリ削除 実行終了
    echo msgbox "処理が完了しました",vbInformation,"処理完了" > %TEMP%\msgbox.vbs & %TEMP%\msgbox.vbs

    REM MsgBox表示用_一時vbs_削除
    timeout /t 0 >nul
    del /Q %TEMP%\msgbox.vbs
) else (
    REM 完了表示
    echo レジストリ削除 中断
    echo msgbox "処理を中断しました",vbExclamation,"処理中断" > %TEMP%\msgbox.vbs & %TEMP%\msgbox.vbs

    REM MsgBox表示用_一時vbs_削除
    timeout /t 0 >nul
    del /Q %TEMP%\msgbox.vbs
)




REM ==============================================
REM bat終了
exit

