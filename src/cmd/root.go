package cmd

import (
	"bytes"
	"fmt"
	"os"
	"parking/src/conf"
	"parking/src/read"
	"parking/src/transfer"
	"parking/src/word"
	"slices"

	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

var (
	cfgFile  string
	output   string
	input    string
	namePool string
	fileName string
	rootCmd  = &cobra.Command{
		Use:   "parking",
		Short: "construct parking attendances file",
		Long:  "construct parking attendances file",
		Run: func(cmd *cobra.Command, args []string) {

			fmt.Printf("all settings : %+v\n", viper.AllSettings())

			fmt.Println("input  \n", viper.GetString("input"))

			fixedMap, tempMap, err := read.ReadData(viper.GetString("input"))

			if err != nil {
				panic(err)
			}

			fmt.Printf("fixed ：%+v\n", fixedMap)
			fmt.Printf("temp : %+v\n", tempMap)

			billData := transfer.CreateBillData(fixedMap, tempMap)
			word.CreateBillDoc(&billData)

			fixedPDList := transfer.FixedTransfer(fixedMap)
			fmt.Printf("fixed page list : %+v \n", fixedPDList)
			fmt.Printf("-------------------------------------")
			tempPDList := transfer.TempTransfer(tempMap)
			fmt.Printf("temp page list : %+v\n", tempPDList)
			fmt.Printf("-------------------------------------")

			list := slices.Concat(fixedPDList, tempPDList)
			fmt.Printf("list : %+v", list)
			word.CreateDocx(list)

		},
	}
)

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
	}
}

func init() {
	cobra.OnInitialize(initConfig)

	rootCmd.PersistentFlags().StringVarP(&cfgFile, "config", "c", "", "config file")
	rootCmd.PersistentFlags().StringVarP(&output, "output", "o", "D:\\work_space\\parking\\output", "output path")
	rootCmd.PersistentFlags().StringVarP(&input, "input", "i", "D:\\work_space\\parking\\input\\input.xlsx", "input file")
	rootCmd.PersistentFlags().StringVarP(&namePool, "name-pool", "n", "D:\\work_space\\parking\\source\\name_pool.log", "name pool")
	rootCmd.PersistentFlags().StringVarP(&fileName, "output-file", "O", "att.docx", "file name")

	viper.BindPFlag("output", rootCmd.PersistentFlags().Lookup("output"))
	viper.BindPFlag("input", rootCmd.PersistentFlags().Lookup("input"))
	viper.BindPFlag("name_pool", rootCmd.PersistentFlags().Lookup("name-pool"))
	viper.BindPFlag("file", rootCmd.PersistentFlags().Lookup("output-file"))

	viper.SetDefault("input", "D:\\work_space\\parking\\input\\input.xlsx")
	viper.SetDefault("output", "D:\\work_space\\parking\\output")
	viper.SetDefault("name_pool", "C:\\Users\\Lenovo\\name_pool.log")
	viper.SetDefault("file", "att.docx")
	viper.SetDefault("temp_8_day", 194)
	viper.SetDefault("temp_12_day", 247)
	viper.SetDefault("fixed_pay", 4580)
	viper.SetDefault("contract_start", "2025-01-01")
	viper.SetDefault("contract_end", "2025-12-31")
}

func initConfig() {

	conf.ReadConfig()

	if cfgFile != "" {

		bs, err := os.ReadFile(cfgFile)

		viper.MergeConfig(bytes.NewReader(bs))

		if err != nil {
			panic(err)
		}
	}

	fmt.Printf("init config  %+v \n", viper.AllSettings())

}
