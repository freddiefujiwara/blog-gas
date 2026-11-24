/-
  JSON Lines resident Test Oracle for blog-gas's translation
  segmentation.

  Protocol (one JSON object per line on stdin, one per line on
  stdout):
    {"op":"segments","markdown":"..."}
      -> {"segments":[{"t":"...","p":true},...],
          "calls":["<joined batch>",...]}
    {"op":"splitLong","text":"...","limit":4000}
      -> {"parts":["...",...]}
-/
import Lean
import BlogGasSpec.Spec

open BlogGasSpec

/-- Minimal JSON string escaping (non-ASCII is emitted raw, which
    is valid JSON and what JSON.parse on the JS side expects). -/
def escJson (s : String) : String :=
  let piece (c : Char) : String :=
    if c == '"' then "\\\""
    else if c == '\\' then "\\\\"
    else if c == '\n' then "\\n"
    else if c == '\r' then "\\r"
    else if c == '\t' then "\\t"
    else if c.toNat < 0x20 then
      let hex := "0123456789abcdef".toList
      "\\u00" ++ String.singleton (hex.get! (c.toNat / 16)) ++
        String.singleton (hex.get! (c.toNat % 16))
    else String.singleton c
  "\"" ++ s.toList.foldl (fun acc c => acc ++ piece c) "" ++ "\""

def segJson (seg : String × Bool) : String :=
  "{\"t\":" ++ escJson seg.1 ++ ",\"p\":" ++
    (if seg.2 then "true" else "false") ++ "}"

def strArrJson (xs : List String) : String :=
  "[" ++ String.intercalate "," (xs.map escJson) ++ "]"

def handle (line : String) : String :=
  match Lean.Json.parse line with
  | .error e => "{\"error\":\"parse: " ++ e ++ "\"}"
  | .ok j =>
    let op := ((j.getObjValD "op").getStr?).toOption.getD ""
    if op == "segments" then
      match (j.getObjValD "markdown").getStr? with
      | .error e => "{\"error\":\"markdown: " ++ e ++ "\"}"
      | .ok md =>
        "{\"segments\":[" ++
          String.intercalate "," ((segments md).map segJson) ++
          "],\"calls\":" ++ strArrJson (calls md) ++ "}"
    else if op == "splitLong" then
      match (j.getObjValD "text").getStr?,
            (j.getObjValD "limit").getNat? with
      | .ok text, .ok limit =>
        "{\"parts\":" ++ strArrJson (splitLong text limit) ++ "}"
      | _, _ => "{\"error\":\"splitLong fields\"}"
    else "{\"error\":\"unknown op\"}"

def main : IO Unit := do
  let stdin ← IO.getStdin
  let stdout ← IO.getStdout
  let mut line ← stdin.getLine
  while !line.isEmpty do
    let trimmed := line.trim
    if !trimmed.isEmpty then
      stdout.putStr (handle trimmed ++ "\n")
      stdout.flush
    line ← stdin.getLine
