import Lake
open Lake DSL

package bloggas_spec where

lean_lib BlogGasSpec where

lean_exe bloggas_oracle where
  root := `Main
