#include <iostream>
#include <string>
#include <vector>
#include <algorithm>
#include <Windows.h>
#include <mfapi.h>
#include <mfidl.h>
#include <mfreadwrite.h>
#include <shlwapi.h>
#include <wincodec.h>
#include <cstdlib>
#include <sstream>
#include <iomanip>
#include <dshow.h>
#include <strmif.h>

// 添加库链接指令
#pragma comment(lib, "mf.lib")
#pragma comment(lib, "mfplat.lib")
#pragma comment(lib, "mfreadwrite.lib")
#pragma comment(lib, "mfuuid.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "strmiids.lib")
#pragma comment(lib, "quartz.lib")

// Helper to release COM objects safely
template <class T>
void SafeRelease(T **ppT)
{
    if (*ppT)
    {
        (*ppT)->Release();
        *ppT = NULL;
    }
}

// 从符号链接中提取VID和PID
std::wstring ExtractVidPid(const std::wstring &symbolicLink)
{
    std::wstring result;

    // 查找 vid_ 和 pid_ 模式
    size_t vidPos = symbolicLink.find(L"vid_");
    size_t pidPos = symbolicLink.find(L"pid_");

    if (vidPos != std::wstring::npos && pidPos != std::wstring::npos)
    {
        // 提取VID (4个字符)
        std::wstring vid = symbolicLink.substr(vidPos + 4, 4);
        // 提取PID (4个字符)
        std::wstring pid = symbolicLink.substr(pidPos + 4, 4);

        result = vid + pid; // 组合VID和PID
    }

    return result;
}

// 根据格式获取文件扩展名
std::wstring GetFileExtension(const GUID &format)
{
    if (format == MFVideoFormat_MJPG)
        return L".jpg";
    else if (format == MFVideoFormat_RGB24)
        return L".bmp";
    else if (format == MFVideoFormat_NV12)
        return L".nv12";
    else if (format == MFVideoFormat_YUY2)
        return L".yuy2";
    else
        return L".raw";
}

// 直接保存原始数据到文件
HRESULT SaveRawDataToFile(const WCHAR *basePath, const GUID &format, UINT32 width, UINT32 height, BYTE *data, DWORD dataSize)
{
    if (basePath == NULL || data == NULL || dataSize == 0)
    {
        std::wcerr << L"SaveRawDataToFile: Invalid parameters" << std::endl;
        return E_INVALIDARG;
    }

    std::wstring extension = GetFileExtension(format);
    std::wstring fullPath = std::wstring(basePath) + extension;

    // std::wcout << L"Attempting to save to: " << fullPath << std::endl;

    HANDLE hFile = CreateFileW(
        fullPath.c_str(),
        GENERIC_WRITE,
        0,
        NULL,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        NULL);

    if (hFile == INVALID_HANDLE_VALUE)
    {
        DWORD error = GetLastError();
        std::wcerr << L"CreateFileW failed. Error: " << error << std::endl;
        return HRESULT_FROM_WIN32(error);
    }

    // std::wcout << L"File created successfully, writing " << dataSize << L" bytes..." << std::endl;

    DWORD bytesWritten = 0;
    BOOL result = WriteFile(hFile, data, dataSize, &bytesWritten, NULL);

    if (!result)
    {
        DWORD error = GetLastError();
        std::wcerr << L"WriteFile failed. Error: " << error << std::endl;
        CloseHandle(hFile);
        return HRESULT_FROM_WIN32(error);
    }

    CloseHandle(hFile);

    if (bytesWritten != dataSize)
    {
        std::wcerr << L"Partial write: wrote " << bytesWritten << L" of " << dataSize << L" bytes" << std::endl;
        return E_FAIL;
    }

    std::wcout << L"Successfully saved " << bytesWritten << L" bytes to " << fullPath << std::endl;
    std::wcout << L"Format: ";

    if (format == MFVideoFormat_MJPG)
        std::wcout << L"MJPG";
    else if (format == MFVideoFormat_RGB24)
        std::wcout << L"RGB24";
    else if (format == MFVideoFormat_NV12)
        std::wcout << L"NV12";
    else if (format == MFVideoFormat_YUY2)
        std::wcout << L"YUY2";
    else
        std::wcout << L"Unknown";

    std::wcout << L", Resolution: " << width << L"x" << height << std::endl;

    return S_OK;
}

void adjust_camera_settings(IMFMediaSource *pSource, IAMVideoProcAmp *pVideoProcAmp = NULL)
{
    if (pSource == NULL)
        return;

    HRESULT vidHr = pSource->QueryInterface(IID_IAMVideoProcAmp, (void **)&pVideoProcAmp);
    if (SUCCEEDED(vidHr))
    {
        // 白平衡
        LONG wbValue = 0, wbFlags = 0;
        if (SUCCEEDED(pVideoProcAmp->Get(VideoProcAmp_WhiteBalance, &wbValue, &wbFlags)))
        {
            pVideoProcAmp->Set(VideoProcAmp_WhiteBalance, wbValue, VideoProcAmp_Flags_Auto);
            std::wcout << L"  Auto white balance enabled" << std::endl;
        }

        // 背光补偿
        LONG backlightValue = 0, backlightFlags = 0;
        if (SUCCEEDED(pVideoProcAmp->Get(VideoProcAmp_BacklightCompensation, &backlightValue, &backlightFlags)))
        {
            pVideoProcAmp->Set(VideoProcAmp_BacklightCompensation, 1, VideoProcAmp_Flags_Manual);
            std::wcout << L"  Backlight compensation enabled" << std::endl;
        }

        // 亮度
        LONG brightnessMin, brightnessMax, brightnessStep, brightnessDefault, brightnessCaps;
        if (SUCCEEDED(pVideoProcAmp->GetRange(VideoProcAmp_Brightness, &brightnessMin, &brightnessMax, &brightnessStep, &brightnessDefault, &brightnessCaps)))
        {
            LONG targetBrightness = brightnessDefault + (brightnessMax - brightnessDefault) / 3;
            if (targetBrightness > brightnessMax)
                targetBrightness = brightnessMax;

            pVideoProcAmp->Set(VideoProcAmp_Brightness, targetBrightness, VideoProcAmp_Flags_Manual);
            std::wcout << L"  Brightness increased to " << targetBrightness
                       << L" (range: " << brightnessMin << L"-" << brightnessMax << L")" << std::endl;
        }

        // 对比度
        LONG contrastMin, contrastMax, contrastStep, contrastDefault, contrastCaps;
        if (SUCCEEDED(pVideoProcAmp->GetRange(VideoProcAmp_Contrast, &contrastMin, &contrastMax, &contrastStep, &contrastDefault, &contrastCaps)))
        {
            LONG targetContrast = contrastDefault + (contrastMax - contrastDefault) / 2;
            if (targetContrast > contrastMax)
                targetContrast = contrastMax;

            pVideoProcAmp->Set(VideoProcAmp_Contrast, targetContrast, VideoProcAmp_Flags_Manual);
            std::wcout << L"  Contrast significantly increased to " << targetContrast
                       << L" (range: " << contrastMin << L"-" << contrastMax << L")" << std::endl;
        }

        // 饱和度
        LONG saturationMin, saturationMax, saturationStep, saturationDefault, saturationCaps;
        if (SUCCEEDED(pVideoProcAmp->GetRange(VideoProcAmp_Saturation, &saturationMin, &saturationMax, &saturationStep, &saturationDefault, &saturationCaps)))
        {
            LONG targetSaturation = saturationDefault + (saturationMax - saturationDefault) / 4;
            if (targetSaturation > saturationMax)
                targetSaturation = saturationMax;

            pVideoProcAmp->Set(VideoProcAmp_Saturation, targetSaturation, VideoProcAmp_Flags_Manual);
            std::wcout << L"  Saturation increased to " << targetSaturation
                       << L" (range: " << saturationMin << L"-" << saturationMax << L")" << std::endl;
        }

        // 锐度
        LONG sharpnessMin, sharpnessMax, sharpnessStep, sharpnessDefault, sharpnessCaps;
        if (SUCCEEDED(pVideoProcAmp->GetRange(VideoProcAmp_Sharpness, &sharpnessMin, &sharpnessMax, &sharpnessStep, &sharpnessDefault, &sharpnessCaps)))
        {
            LONG targetSharpness = sharpnessDefault + (sharpnessMax - sharpnessDefault) / 2;
            if (targetSharpness > sharpnessMax)
                targetSharpness = sharpnessMax;

            pVideoProcAmp->Set(VideoProcAmp_Sharpness, targetSharpness, VideoProcAmp_Flags_Manual);
            std::wcout << L"  Sharpness increased to " << targetSharpness
                       << L" (range: " << sharpnessMin << L"-" << sharpnessMax << L")" << std::endl;
        }

        // 尝试调整伽马值
        LONG gammaMin, gammaMax, gammaStep, gammaDefault, gammaCaps;
        if (SUCCEEDED(pVideoProcAmp->GetRange(VideoProcAmp_Gamma, &gammaMin, &gammaMax, &gammaStep, &gammaDefault, &gammaCaps)))
        {
            LONG targetGamma = gammaDefault - (gammaDefault - gammaMin) / 4;
            if (targetGamma < gammaMin)
                targetGamma = gammaMin;

            pVideoProcAmp->Set(VideoProcAmp_Gamma, targetGamma, VideoProcAmp_Flags_Manual);
            std::wcout << L"  Gamma adjusted to " << targetGamma
                       << L" (range: " << gammaMin << L"-" << gammaMax << L")" << std::endl;
        }
    }

    // 增加等待时间让所有设置充分生效
    std::wcout << L"Waiting for camera settings to take effect..." << std::endl;
    // Sleep(5000);
}

int main(int argc, char *argv[])
{
    if (argc < 5)
    {
        std::cerr << "Usage: console-camera.exe <output_base_path> <vid:pid> <width> <height>" << std::endl;
        std::cerr << "Example: console-camera.exe output/capture 32e4:9415 3840 2160" << std::endl;
        std::cerr << "Output will be saved as output/capture.jpg (or .nv12, .yuy2, etc.)" << std::endl;
        return EXIT_FAILURE;
    }

    int wargc;
    wchar_t **wargv = CommandLineToArgvW(GetCommandLineW(), &wargc);

    if (wargv == NULL || wargc < 5)
    {
        std::wcerr << L"Usage: console-camera.exe <output_base_path> <vid:pid> <width> <height>" << std::endl;
        return EXIT_FAILURE;
    }

    const WCHAR *outputBasePath = wargv[1];
    const WCHAR *vidPidParam = wargv[2];
    UINT32 targetWidth = static_cast<UINT32>(std::wcstoul(wargv[3], nullptr, 10));
    UINT32 targetHeight = static_cast<UINT32>(std::wcstoul(wargv[4], nullptr, 10));

    // 解析 VID:PID 格式
    std::wstring vidPidStr(vidPidParam);
    std::wstring targetVidPid;

    size_t colonPos = vidPidStr.find(L':');
    if (colonPos != std::wstring::npos && colonPos > 0 && colonPos < vidPidStr.length() - 1)
    {
        std::wstring vid = vidPidStr.substr(0, colonPos);
        std::wstring pid = vidPidStr.substr(colonPos + 1);
        targetVidPid = vid + pid;
        std::wcout << L"Searching for device with VID:PID = " << vid << L":" << pid << std::endl;
    }
    else
    {
        std::wcerr << L"Error: Invalid VID:PID format. Expected format: 32e4:9415" << std::endl;
        LocalFree(wargv);
        return EXIT_FAILURE;
    }

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr))
    {
        std::wcerr << L"Error: CoInitializeEx failed." << std::endl;
        LocalFree(wargv);
        return EXIT_FAILURE;
    }

    hr = MFStartup(MF_VERSION);
    if (FAILED(hr))
    {
        std::wcerr << L"Error: MFStartup failed." << std::endl;
        CoUninitialize();
        LocalFree(wargv);
        return EXIT_FAILURE;
    }

    IMFAttributes *pAttributes = NULL;
    IMFActivate **ppDevices = NULL;
    UINT32 count = 0;
    IMFActivate *pTargetDevice = NULL;

    hr = MFCreateAttributes(&pAttributes, 1);
    if (SUCCEEDED(hr))
    {
        hr = pAttributes->SetGUID(MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE, MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_GUID);
    }
    if (SUCCEEDED(hr))
    {
        hr = MFEnumDeviceSources(pAttributes, &ppDevices, &count);
    }

    // Find the requested device by VID:PID
    if (SUCCEEDED(hr))
    {
        for (UINT32 i = 0; i < count; i++)
        {
            WCHAR *symbolicLink = NULL;
            WCHAR *friendlyName = NULL;

            HRESULT hrSymLink = ppDevices[i]->GetAllocatedString(MF_DEVSOURCE_ATTRIBUTE_SOURCE_TYPE_VIDCAP_SYMBOLIC_LINK, &symbolicLink, NULL);
            HRESULT hrName = ppDevices[i]->GetAllocatedString(MF_DEVSOURCE_ATTRIBUTE_FRIENDLY_NAME, &friendlyName, NULL);

            std::wstring vidPid = ExtractVidPid(symbolicLink ? symbolicLink : L"");

            if (!vidPid.empty())
            {
                std::wstring vidPidLower = vidPid;
                std::wstring targetLower = targetVidPid;
                std::transform(vidPidLower.begin(), vidPidLower.end(), vidPidLower.begin(), ::towlower);
                std::transform(targetLower.begin(), targetLower.end(), targetLower.begin(), ::towlower);

                if (vidPidLower == targetLower)
                {
                    pTargetDevice = ppDevices[i];
                    pTargetDevice->AddRef();
                    std::wcout << L"Found device: " << (friendlyName ? friendlyName : L"Unknown") << L" (" << vidPid << L")" << std::endl;
                    break;
                }
            }

            CoTaskMemFree(symbolicLink);
            CoTaskMemFree(friendlyName);
            SafeRelease(&ppDevices[i]);
        }
        CoTaskMemFree(ppDevices);
    }

    if (pTargetDevice == NULL)
    {
        std::wcerr << L"Error: Could not find camera device with VID:PID " << targetVidPid << std::endl;
        hr = E_FAIL;
    }

    IMFMediaSource *pSource = NULL;
    IMFSourceReader *pReader = NULL;
    GUID selectedFormat = GUID_NULL;

    if (SUCCEEDED(hr))
    {
        hr = pTargetDevice->ActivateObject(IID_PPV_ARGS(&pSource));
        if (FAILED(hr))
        {
            std::wcerr << L"Failed to activate device. HRESULT = 0x" << std::hex << hr << std::endl;
        }
    }

    if (SUCCEEDED(hr))
    {
        // 创建 SourceReader 时添加属性配置
        IMFAttributes *pReaderAttributes = NULL;
        hr = MFCreateAttributes(&pReaderAttributes, 2);

        if (SUCCEEDED(hr))
        {
            // 禁用硬件变换以避免兼容性问题
            hr = pReaderAttributes->SetUINT32(MF_SOURCE_READER_DISABLE_DXVA, TRUE);
        }

        if (SUCCEEDED(hr))
        {
            // 启用视频处理
            hr = pReaderAttributes->SetUINT32(MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING, TRUE);
        }

        if (SUCCEEDED(hr))
        {
            hr = MFCreateSourceReaderFromMediaSource(pSource, pReaderAttributes, &pReader);
            if (FAILED(hr))
            {
                std::wcerr << L"Failed to create SourceReader. HRESULT = 0x" << std::hex << hr << std::endl;
            }
        }

        SafeRelease(&pReaderAttributes);
    }

    // 查找匹配的分辨率并使用原生格式，优先使用 MJPG
    if (SUCCEEDED(hr))
    {
        IMFMediaType *pBestType = NULL;
        DWORD typeIndex = 0;
        bool foundSuitableType = false;

        // 格式优先级：MJPG > RGB24 > YUY2 > NV12
        std::vector<GUID> preferredFormats = {
            MFVideoFormat_MJPG,
            MFVideoFormat_RGB24,
            MFVideoFormat_YUY2,
            MFVideoFormat_NV12};

        // 按优先级顺序查找匹配的格式和分辨率
        for (const auto &preferredFormat : preferredFormats)
        {
            typeIndex = 0;
            while (SUCCEEDED(pReader->GetNativeMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM, typeIndex, &pBestType)))
            {
                GUID subtype;
                UINT32 width, height;

                if (SUCCEEDED(pBestType->GetGUID(MF_MT_SUBTYPE, &subtype)) &&
                    SUCCEEDED(MFGetAttributeSize(pBestType, MF_MT_FRAME_SIZE, &width, &height)))
                {
                    // 检查分辨率和格式匹配
                    if (width == targetWidth && height == targetHeight && subtype == preferredFormat)
                    {
                        std::wcout << L"Found matching format: " << width << L"x" << height;

                        if (subtype == MFVideoFormat_MJPG)
                        {
                            std::wcout << L" (MJPG)";
                        }
                        else if (subtype == MFVideoFormat_RGB24)
                            std::wcout << L" (RGB24)";
                        else if (subtype == MFVideoFormat_YUY2)
                            std::wcout << L" (YUY2)";
                        else if (subtype == MFVideoFormat_NV12)
                            std::wcout << L" (NV12)";

                        std::wcout << L" - Attempting to set..." << std::endl;

                        // 方法1: 直接使用原始媒体类型
                        hr = pReader->SetCurrentMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM, NULL, pBestType);

                        if (FAILED(hr) && subtype == MFVideoFormat_MJPG)
                        {
                            std::wcout << L"  Direct MJPG failed, trying simplified MJPG type..." << std::endl;

                            // 方法2: 创建一个简化的 MJPG 媒体类型
                            IMFMediaType *pSimpleMJPG = NULL;
                            HRESULT createHr = MFCreateMediaType(&pSimpleMJPG);

                            if (SUCCEEDED(createHr))
                            {
                                createHr = pSimpleMJPG->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video);
                                if (SUCCEEDED(createHr))
                                    createHr = pSimpleMJPG->SetGUID(MF_MT_SUBTYPE, MFVideoFormat_MJPG);
                                if (SUCCEEDED(createHr))
                                    createHr = MFSetAttributeSize(pSimpleMJPG, MF_MT_FRAME_SIZE, width, height);

                                if (SUCCEEDED(createHr))
                                {
                                    hr = pReader->SetCurrentMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM, NULL, pSimpleMJPG);
                                    std::wcout << L"  Simplified MJPG result: 0x" << std::hex << hr << std::endl;
                                }

                                SafeRelease(&pSimpleMJPG);
                            }
                        }

                        if (SUCCEEDED(hr))
                        {
                            selectedFormat = subtype;
                            foundSuitableType = true;
                            std::wcout << L"Media type set successfully!" << std::endl;
                        }
                        else
                        {
                            std::wcerr << L"Failed to set media type. HRESULT = 0x" << std::hex << hr << L", trying next format..." << std::endl;
                            // 继续尝试其他格式
                            hr = S_OK; // 重置 hr 以继续尝试其他格式
                        }

                        SafeRelease(&pBestType);

                        // 如果成功设置了媒体类型，立即退出所有循环
                        if (foundSuitableType)
                        {
                            goto format_selection_complete;
                        }

                        break;
                    }
                }

                SafeRelease(&pBestType);
                typeIndex++;
            }

            if (foundSuitableType)
                break;
        }

    format_selection_complete:

        if (!foundSuitableType)
        {
            std::wcerr << L"Error: Device does not support resolution " << targetWidth << L"x" << targetHeight << std::endl;
            hr = E_FAIL;
        }
    }

    // 配置摄像头参数以提高图像质量
    if (SUCCEEDED(hr) && pSource)
    {
        std::wcout << L"Configuring camera parameters for enhanced quality..." << std::endl;

        // 获取摄像头控制接口
        IAMCameraControl *pCameraControl = NULL;
        IAMVideoProcAmp *pVideoProcAmp = NULL;

        HRESULT camHr = pSource->QueryInterface(IID_IAMCameraControl, (void **)&pCameraControl);
        if (SUCCEEDED(camHr))
        {
            // 设置焦点为自动
            LONG focusValue = 0, focusFlags = 0;
            if (SUCCEEDED(pCameraControl->Get(CameraControl_Focus, &focusValue, &focusFlags)))
            {
                pCameraControl->Set(CameraControl_Focus, focusValue, CameraControl_Flags_Auto);
                std::wcout << L"  Auto focus enabled" << std::endl;
            }

            // 偏向更长曝光时间
            LONG exposureMin, exposureMax, exposureStep, exposureDefault, exposureCaps;
            if (SUCCEEDED(pCameraControl->GetRange(CameraControl_Exposure, &exposureMin, &exposureMax, &exposureStep, &exposureDefault, &exposureCaps)))
            {
                if (exposureCaps & CameraControl_Flags_Auto)
                {
                    pCameraControl->Set(CameraControl_Exposure, exposureDefault, CameraControl_Flags_Auto);
                    std::wcout << L"  Auto exposure enabled" << std::endl;
                }
                else
                {
                    // 手动设置稍长的曝光时间以提高暗部细节
                    LONG targetExposure = exposureDefault + (exposureMax - exposureDefault) / 4;
                    if (targetExposure > exposureMax)
                        targetExposure = exposureMax;
                    pCameraControl->Set(CameraControl_Exposure, targetExposure, CameraControl_Flags_Manual);
                    std::wcout << L"  Manual exposure set to " << targetExposure << std::endl;
                }
            }

            SafeRelease(&pCameraControl);
        }

        // 效果微弱，弃用
        adjust_camera_settings(pSource);
        SafeRelease(&pVideoProcAmp);
    }

    // 捕获帧并保存原始数据
    IMFSample *pSample = NULL;
    if (SUCCEEDED(hr))
    {
        // std::wcout << L"Capturing frame..." << std::endl;

        // 等待设备准备好
        Sleep(500);

        // 重试机制来获取第一帧
        const int MAX_RETRIES = 10;
        int retryCount = 0;

        while (retryCount < MAX_RETRIES && SUCCEEDED(hr))
        {
            DWORD streamFlags = 0;
            LONGLONG timestamp = 0;

            hr = pReader->ReadSample(
                MF_SOURCE_READER_FIRST_VIDEO_STREAM,
                0,
                NULL,
                &streamFlags,
                &timestamp,
                &pSample);

            if (FAILED(hr))
            {
                std::wcerr << L"ReadSample failed. HRESULT = 0x" << std::hex << hr << std::endl;
                break;
            }
            else if (pSample != NULL)
            {
                // std::wcout << L"Frame captured successfully!" << std::endl;
                break;
            }
            else
            {
                // 摄像头预热中，继续重试
                retryCount++;
                if (retryCount < MAX_RETRIES)
                {
                    Sleep(200);
                }
            }
        }

        if (pSample == NULL && SUCCEEDED(hr))
        {
            std::wcerr << L"Failed to get sample after " << MAX_RETRIES << L" attempts." << std::endl;
            hr = E_FAIL;
        }
    }

    if (SUCCEEDED(hr) && pSample)
    {
        // std::wcout << L"Processing captured frame..." << std::endl;

        IMFMediaBuffer *pBuffer = NULL;
        hr = pSample->ConvertToContiguousBuffer(&pBuffer);

        if (SUCCEEDED(hr))
        {
            BYTE *pData = NULL;
            DWORD dataSize = 0;
            DWORD maxLength = 0;

            hr = pBuffer->Lock(&pData, &maxLength, &dataSize);

            if (SUCCEEDED(hr) && pData != NULL && dataSize > 0)
            {
                // std::wcout << L"Buffer locked successfully. Data size: " << dataSize << L" bytes" << std::endl;

                hr = SaveRawDataToFile(outputBasePath, selectedFormat, targetWidth, targetHeight, pData, dataSize);

                pBuffer->Unlock();
            }
            else
            {
                std::wcerr << L"Failed to lock buffer or invalid data." << std::endl;
                hr = E_FAIL;
            }
        }
        SafeRelease(&pBuffer);
    }

    if (FAILED(hr))
    {
        std::wcerr << L"Error: Failed to capture or save frame. HRESULT = 0x" << std::hex << hr << std::endl;
    }

    SafeRelease(&pSample);
    SafeRelease(&pReader);
    SafeRelease(&pSource);
    SafeRelease(&pTargetDevice);
    SafeRelease(&pAttributes);

    MFShutdown();
    CoUninitialize();
    LocalFree(wargv);

    return SUCCEEDED(hr) ? EXIT_SUCCESS : EXIT_FAILURE;
}