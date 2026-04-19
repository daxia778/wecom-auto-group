package main

import (
	"embed"
	"log"
	"os"

	"github.com/wailsapp/wails/v2"
	"github.com/wailsapp/wails/v2/pkg/options"
	"github.com/wailsapp/wails/v2/pkg/options/assetserver"
	"github.com/wailsapp/wails/v2/pkg/options/windows"
)

//go:embed all:frontend/dist
var assets embed.FS

func main() {
	// --test 模式: 跳过 GUI，直接测试 API
	for _, arg := range os.Args[1:] {
		if arg == "--test" || arg == "-test" {
			TestWeComAPI()
			return
		}
		if arg == "--diag" || arg == "-diag" {
			TestWindowOCR()
			return
		}
		if arg == "--privacy-test" || arg == "-privacy-test" {
			TestPrivacyFlow()
			return
		}
		if arg == "--spy-windows" || arg == "-spy-windows" {
			TestSpyWindows()
			return
		}
		if arg == "--click-test" || arg == "-click-test" {
			TestBackendClick()
			return
		}
		if arg == "--screenshot-test" || arg == "-screenshot-test" {
			TestScreenshotCapabilities()
			return
		}
		if arg == "--interactive" || arg == "-interactive" {
			InteractivePrivacyTest()
			return
		}
		if arg == "--group-test" || arg == "-group-test" {
			TestGroupCreation()
			return
		}
	}

	app := NewApp()

	err := wails.Run(&options.App{
		Title:     "WeCom 自动建群 v2.0",
		Width:     900,
		Height:    650,
		MinWidth:  800,
		MinHeight: 600,
		AssetServer: &assetserver.Options{
			Assets: assets,
		},
		OnStartup: app.startup,
		Bind: []interface{}{
			app,
		},
		Windows: &windows.Options{
			WebviewIsTransparent: true,
			WindowIsTranslucent:  false,
			DisableWindowIcon:    false,
		},
	})

	if err != nil {
		log.Fatal(err)
	}
}
