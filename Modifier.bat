@echo off
color 02
title Modifier

@REM 检查是否以管理员身份运行批处理
net session >nul 2>&1
if %errorlevel% neq 0 (
    color 04
    echo 请在打开时右键文件, 使用管理员身份运行批处理!
    pause
    goto x
)

@REM 工作目录检测
:FindInstallDirectory
echo 正在寻找工作目录
if exist "%ProgramFiles(x86)%\Kingsoft\WPS Office\" (
    set "ActiveDirectory=%ProgramFiles(x86)%\Kingsoft\WPS Office"
    goto FindResourceDirectory
) else if exist ".\WPS Office\" (
    set "ActiveDirectory=.\WPS Office"
    goto FindResourceDirectory
) else (
    goto Error
)

@REM 检查资源目录
:FindResourceDirectory
if not exist ".\releases\" (
    mkdir ".\releases\"
    explorer ".\releases\"
    goto ReleaseFileTips
) else (
    goto ReleaseFileChecker
)

@REM 检查资源文件
:ReleaseFileChecker
echo 正在检查资源文件
if not exist ".\releases\logo.png" (
    goto ReleaseFileTips
) else (
    if not exist ".\releases\qrcode.png" (
        goto ReleaseFileTips
    ) else (
        if not exist ".\releases\qrcode.jpg" (
            goto ReleaseFileTips
        ) else (
            if not exist ".\releases\ent-oem.png" (
                goto ReleaseFileTips
            ) else (
                if not exist ".\releases\ent-def.png" (
                    goto ReleaseFileTips
                ) else (
                    goto FindSoftwareRootDirectory
                )
            )
        )
    )
)

@REM 软件根目录检测
:FindSoftwareRootDirectory
color 02
@REM 注: 若修改的版本不是当前批处理中的版本号, 请修改成当前安装的版本号
set "version=12.8.2.18205"
echo 正在检查软件根目录
if exist "%ActiveDirectory%\%version%" (
    echo 找到版本号为 %version% 的文件夹
    set "FullActiveDirectory=%ActiveDirectory%\%version%"
    goto FindAndModifyOEMDirectory
) else (
    echo 目标查找目录 %FullActiveDirectory%
    echo 未找到版本号为 %version% 的文件夹
    echo 请检查变量版本号与安装的版本是否匹配!
    pause
    goto x
)

@REM 查找与替换
:FindAndModifyOEMDirectory
@REM Logo部分
if exist "%FullActiveDirectory%\oem\companylogo.png" (
    echo 正在替换 ~\oem\companylogo.png
    copy /y ".\releases\logo.png" "%FullActiveDirectory%\oem\companylogo.png"
)

if exist "%FullActiveDirectory%\utility\backup\OemFile\oem\companylogo.png" (
    echo 正在替换 ~\utility\backup\OemFile\oem\companylogo.png
    copy /y ".\releases\logo.png" "%FullActiveDirectory%\utility\backup\OemFile\oem\companylogo.png"
)

@REM QRCode
@REM PNG
if exist "%FullActiveDirectory%\oem\wechat_dq_image.png" (
    echo 正在替换 ~\oem\wechat_dq_image.png
    copy /y ".\releases\qrcode.png" "%FullActiveDirectory%\oem\wechat_dq_image.png"
)

if exist "%FullActiveDirectory%\utility\backup\OemFile\oem\wechat_dq_image.png" (
    echo 正在替换 ~\utility\backup\OemFile\oem\wechat_dq_image.png
    copy /y ".\releases\qrcode.png" "%FullActiveDirectory%\utility\backup\OemFile\oem\wechat_dq_image.png"
)

@REM JPG
if exist "%FullActiveDirectory%\utility\backup\OemFile\office6\cfgs\oeminfo\qrcode.jpg" (
    echo 正在替换 ~\utility\backup\OemFile\office6\cfgs\oeminfo\qrcode.jpg
    copy /y ".\releases\qrcode.jpg" "%FullActiveDirectory%\utility\backup\OemFile\office6\cfgs\oeminfo\qrcode.jpg"
)

if exist "%FullActiveDirectory%\office6\cfgs\oeminfo\qrcode.jpg" (
    echo 正在替换 ~\office6\cfgs\oeminfo\qrcode.jpg
    copy /y ".\releases\qrcode.jpg" "%FullActiveDirectory%\office6\cfgs\oeminfo\qrcode.jpg"
)

@REM Ent部分
@REM OEM
if exist "%FullActiveDirectory%\office6\mui\zh_CN\resource\splash\hdpi\2x\ent_background_2023_oem.png" (
    echo 正在替换 ~\office6\mui\zh_CN\resource\splash\hdpi\2x\ent_background_2023_oem.png
    copy /y ".\releases\ent-oem.png" "%FullActiveDirectory%\office6\mui\zh_CN\resource\splash\hdpi\2x\ent_background_2023_oem.png"
)

if exist "%FullActiveDirectory%\office6\mui\zh_CN\resource\splash\ent_background_2023_oem.png" (
    echo 正在替换 ~\office6\mui\zh_CN\resource\splash\ent_background_2023_oem.png
    copy /y ".\releases\ent-oem.png" "%FullActiveDirectory%\office6\mui\zh_CN\resource\splash\ent_background_2023_oem.png"
)
@REM Default
if exist "%FullActiveDirectory%\office6\mui\zh_CN\resource\splash\hdpi\2x\ent_background_2023_default.png" (
    echo 正在替换 ~\office6\mui\zh_CN\resource\splash\hdpi\2x\ent_background_2023_oem.png
    copy /y ".\releases\ent-def.png" "%FullActiveDirectory%\office6\mui\zh_CN\resource\splash\hdpi\2x\ent_background_2023_default.png"
)

if exist "%FullActiveDirectory%\office6\mui\zh_CN\resource\splash\ent_background_2023_default.png" (
    echo 正在替换 ~\office6\mui\zh_CN\resource\splash\ent_background_2023_oem.png
    copy /y ".\releases\ent-def.png" "%FullActiveDirectory%\office6\mui\zh_CN\resource\splash\ent_background_2023_default.png"
)

echo 运行结束
pause
goto x

:ReleaseFileTips
color 06
echo 请在 release 文件中放置以下文件:
echo 1. logo.png
echo 该文件用于显示启动时的Logo, 建议使用透明背景
echo 2. qrcode.png
echo 3. qrcode.jpg
echo 两个图片文件必须存在
echo 4. ent-oem.png
echo 5. ent-def.png
echo 该文件用于显示在启动界面
echo 放置完成后请按任意键继续
echo 再次检测文件
echo 若要关闭按Ctrl+C结束运行
echo.
pause
goto ReleaseFileChecker

:Error
color 04
echo 发生错误, 请检查工作目录是否存在
echo 若更改了安装位置, 请将批处理连同releases文件夹放置于安装目录 Kingsoft下
echo 该文件夹中必须包含 WPS Office 文件夹
pause
goto x

:x
exit