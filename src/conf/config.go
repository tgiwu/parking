package conf

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)
//默认配置文件在用户目录
func ReadConfig() {

	home, err:= os.UserHomeDir()

	cobra.CheckErr(err)

	viper.AddConfigPath(home)
	viper.SetConfigName("config_common.yaml")

	viper.SetConfigType("yaml")

	if err := viper.ReadInConfig(); err != nil {
		if _, ok := err.(viper.ConfigFileNotFoundError); ok {
			fmt.Println("not find " + err.Error())
		} else {
			fmt.Printf("read config file err, %v \n", err)
		}
	}

	// configName := "config"
	// switch runtime.GOOS {
	// case "windows":
	// 	configName += "_win"
	// case "linux":
	// 	configName += "_lin"
	// case "darwin":
	// 	configName += "_lin"
	// default:
	// 	fmt.Println("unsupport os ", runtime.GOOS)
	// }

	// bs, err := os.ReadFile(filepath.Join(CONFIG_PATH, configName+".yaml"))

	// vip.MergeConfig(bytes.NewReader(bs))

	// if err != nil {
	// 	panic(err)
	// }

	// err = vip.Unmarshal(&MConf)
	// if err != nil {
	// 	panic(err)
	// }
	// viper.AutomaticEnv()
}
