/-
  Executable specification of blog-gas's translation-time markdown
  handling (src/Code.js): `splitMarkdownSegments_`, `splitLongText_`
  and the LanguageApp batch planning inside `translateMarkdown_`.

  Written as an independent re-implementation of the documented
  semantics, to serve as the Test Oracle in differential testing.
  Everything works on `List Char`; for the BMP characters the
  differential harness generates, a JS string's UTF-16 length and
  its character count agree, so JS `length`/`slice`/`lastIndexOf`
  map onto list length/take/drop directly.
-/
namespace BlogGasSpec

/-- JS `\s` (the whitespace class used by the regexes and by
    String.prototype.trim), as code points. -/
def isWs (c : Char) : Bool :=
  let n := c.toNat
  (0x09 ≤ n && n ≤ 0x0D) || n == 0x20 || n == 0xA0 || n == 0x1680 ||
  (0x2000 ≤ n && n ≤ 0x200A) || n == 0x2028 || n == 0x2029 ||
  n == 0x202F || n == 0x205F || n == 0x3000 || n == 0xFEFF

def trimStartC (s : List Char) : List Char := s.dropWhile isWs

def trimEndC (s : List Char) : List Char := (s.reverse.dropWhile isWs).reverse

def trimC (s : List Char) : List Char := trimEndC (trimStartC s)

def isWsAt (s : List Char) (n : Nat) : Bool :=
  match s[n]? with
  | some c => isWs c
  | none => false

/-! ## splitLongText_ -/

/-- Last index inside the window at which a sentence boundary sits:
    one of 。！？、 or the `.` of a `. ` pair. Mirrors the JS
    `Math.max(window.lastIndexOf(...), ...)` over the window. -/
def lastBoundaryGo : Nat → Option Nat → List Char → Option Nat
  | _, acc, [] => acc
  | i, acc, c :: t =>
    let hit := c == '。' || c == '！' || c == '？' || c == '、' ||
      (c == '.' && t.head? == some ' ')
    lastBoundaryGo (i + 1) (if hit then some i else acc) t

def lastBoundary (w : List Char) : Option Nat := lastBoundaryGo 0 none w

/-- The cut position for one split step: just past the last
    sentence boundary when it lies at or beyond 30% of the limit,
    the hard limit otherwise. `(limit * 3) / 10` is
    `Math.floor(limit * 0.3)` for the limits involved. -/
def chooseCut (limit : Nat) (window : List Char) : Nat :=
  match lastBoundary window with
  | some i => if i >= (limit * 3) / 10 then i + 1 else limit
  | none => limit

/-- Fuel-bounded worker for `splitLongText`; fuel counts down in
    step with the remaining length (each step drops at least one
    character when `limit ≥ 1`). The empty input is its own case,
    as in the JS: no parts at all. -/
def splitLongGo (limit : Nat) : Nat → List Char → List (List Char)
  | _, [] => []
  | 0, rest => [rest]
  | fuel + 1, rest =>
    if limit < 1 then [rest]
    else if rest.length ≤ limit then [rest]
    else
      let cut := chooseCut limit (rest.take limit)
      (rest.take cut) :: splitLongGo limit fuel (rest.drop cut)

/-- `splitLongText_(text, limit)`: parts rejoin with no separator,
    each part at most `limit` long (for `limit ≥ 1`); a limit below
    1 means "no split". -/
def splitLongText (limit : Nat) (text : List Char) : List (List Char) :=
  splitLongGo limit text.length text

/-! ## Paragraph tokenisation: String.split(/(\n{2,})/) -/

def nlRun (s : List Char) : Nat := (s.takeWhile (· == '\n')).length

/-- Gather one text block: everything up to (not including) the
    next run of two or more newlines. Single newlines stay inside
    the block. -/
def gatherText : List Char → List Char × List Char
  | [] => ([], [])
  | s@(c :: t) =>
    if c == '\n' && nlRun s ≥ 2 then ([], s)
    else let (a, b) := gatherText t; (c :: a, b)

/-- Alternating text blocks and separator blocks (`\n{2,}`),
    matching the JS split with a captured separator. A separator
    block is flagged `true`. -/
def tokenizeGo : Nat → List Char → List (List Char × Bool)
  | 0, s => if s.isEmpty then [] else [(s, false)]
  | fuel + 1, s =>
    if s.isEmpty then []
    else if s.head? == some '\n' && nlRun s ≥ 2 then
      (s.take (nlRun s), true) :: tokenizeGo fuel (s.drop (nlRun s))
    else
      let (txt, rest) := gatherText s
      (txt, false) :: tokenizeGo fuel rest

def tokenize (s : List Char) : List (List Char × Bool) :=
  tokenizeGo (s.length + 1) s

/-! ## Lines and fences -/

/-- String.split('\n'): empty pieces are kept. -/
def splitLines : List Char → List (List Char)
  | [] => [[]]
  | '\n' :: t => [] :: splitLines t
  | c :: t => match splitLines t with
              | [] => [[c]]
              | l :: ls => (c :: l) :: ls

/-- A line whose first non-whitespace characters are three
    backticks toggles a code fence (the JS checks
    `line.trimStart().startsWith('```')`). -/
def isFenceLine (line : List Char) : Bool :=
  (trimStartC line).take 3 == ['`', '`', '`']

/-- Scan a block's lines for fences: returns (block is touched by
    a fence, fence state after the block). A block entered inside
    a fence counts as touched from the start. -/
def scanFence (lines : List (List Char)) (isOpen : Bool) : Bool × Bool :=
  lines.foldl
    (fun st line => if isFenceLine line then (true, !st.2) else st)
    (isOpen, isOpen)

/-! ## Whole-block / whole-line protection predicates -/

/-- `[[...]]`-style markers: opening pair, a body with no closing
    bracket, closing pair at the very end. Used for `[[...]]` and
    the full-width `［［...］］`. -/
def isMarker (t : List Char) (o cl : Char) : Bool :=
  t.length ≥ 4 && t.take 2 == [o, o] &&
  t.drop (t.length - 2) == [cl, cl] &&
  !((t.drop 2).take (t.length - 4)).any (· == cl)

/-- A bare URL alone: `^https?:\/\/\S+$`. -/
def isBareUrl (t : List Char) : Bool :=
  let rest :=
    if t.take 8 == "https://".toList then some (t.drop 8)
    else if t.take 7 == "http://".toList then some (t.drop 7)
    else none
  match rest with
  | some r => !r.isEmpty && r.all (fun c => !isWs c)
  | none => false

/-- Table separator rows: `/^\|?[\s:|-]+\|$/`. -/
def isTableSep (t : List Char) : Bool :=
  let inClass (c : Char) : Bool := isWs c || c == ':' || c == '|' || c == '-'
  let check (s : List Char) : Bool :=
    s.length ≥ 2 && s.getLast? == some '|' &&
    !(s.take (s.length - 1)).isEmpty && (s.take (s.length - 1)).all inClass
  check t || (t.head? == some '|' && check t.tail)

def blockProtected (part : List Char) : Bool :=
  let t := trimC part
  t.head? == some '<' || isMarker t '[' ']' || isMarker t '［' '］' ||
  isBareUrl t

def lineProtected (line : List Char) : Bool :=
  let t := trimC line
  t.head? == some '<' || isMarker t '[' ']' || isMarker t '［' '］' ||
  isBareUrl t || isTableSep t

/-! ## Structural line prefixes (LINE_PREFIX_RE_) -/

def wsRun (s : List Char) : Nat := (s.takeWhile isWs).length

def digitRun (s : List Char) : Nat := (s.takeWhile (fun c => c.isDigit)).length

/-- Length of the structural prefix of a prose line: a heading's
    `#…# `, a list item's marker, or a blockquote's `> `. The JS
    regex is `/^(#{1,6}\s+|\s*(?:[-*+]|\d+\.)\s+|\s*>\s?)/`; the
    alternatives are tried in that order, and a failed list match
    still falls through to the blockquote alternative. Returns 0
    when there is no prefix. -/
def matchPrefix (line : List Char) : Nat :=
  let k := (line.takeWhile (· == '#')).length
  if 1 ≤ k && k ≤ 6 && isWsAt line k then k + wsRun (line.drop k)
  else
    let i := wsRun line
    let rest := line.drop i
    let bulletLen : Option Nat :=
      match rest with
      | c :: _ =>
        if c == '-' || c == '*' || c == '+' then some 1
        else
          let d := digitRun rest
          if d ≥ 1 && rest[d]? == some '.' then some (d + 1) else none
      | [] => none
    let quoteLen : Nat :=
      if rest.head? == some '>' then
        i + 1 + (if isWsAt rest 1 then 1 else 0)
      else 0
    match bulletLen with
    | some bl =>
      let w := wsRun (rest.drop bl)
      if w ≥ 1 then i + bl + w else quoteLen
    | none => quoteLen

/-! ## Inline protection (INLINE_PROTECTED_RE_) -/

def runLen (s : List Char) (c : Char) : Nat := (s.takeWhile (· == c)).length

def nonRunLen (s : List Char) (c : Char) : Nat :=
  (s.takeWhile (· != c)).length

/-- Inline code spans: `` `+[^`]*`+ `` with the JS regex's
    backtracking — the opening backtick run is tried longest
    first, then shortened; the middle may not contain a backtick,
    so the only viable middle is the maximal non-backtick run. -/
def matchCodeSpan (s : List Char) : Option Nat :=
  if s.head? != some '`' then none
  else
    let a := runLen s '`'
    (List.range a).reverse.findSome? (fun idx =>
      let j := idx + 1
      let b := nonRunLen (s.drop j) '`'
      if s[j + b]? == some '`' then
        some (j + b + runLen (s.drop (j + b)) '`')
      else none)

/-- Links and images: `(!?)\[([^\]]*)\]\(([^)]*)\)`. The text
    between the brackets is prose; the brackets, the optional
    `!` and the destination are protected. Returns the match
    length and the pieces it produces. -/
def matchLink (s : List Char) : Option (Nat × List (List Char × Bool)) :=
  let start : Option (Nat × List Char) :=
    if s.head? == some '!' then
      if (s.drop 1).head? == some '[' then some (1, s.drop 1) else none
    else if s.head? == some '[' then some (0, s)
    else none
  match start with
  | none => none
  | some (bangLen, afterBang) =>
    let inner := afterBang.drop 1
    let textLen := nonRunLen inner ']'
    let afterText := inner.drop textLen
    if afterText.head? != some ']' then none
    else
      let afterBracket := afterText.drop 1
      if afterBracket.head? != some '(' then none
      else
        let destFull := afterBracket.drop 1
        let destLen := nonRunLen destFull ')'
        let afterDest := destFull.drop destLen
        if afterDest.head? != some ')' then none
        else
          let bangStr : List Char := if bangLen == 1 then ['!'] else []
          let pieces : List (List Char × Bool) :=
            [(bangStr ++ ['['], false)] ++
            (if textLen > 0 then [(inner.take textLen, true)] else []) ++
            [([']', '('] ++ destFull.take destLen ++ [')'], false)]
          some (bangLen + textLen + destLen + 4, pieces)

/-- Bare URLs inside prose: `https?:\/\/\S+`. The scheme and the
    run after it are all non-whitespace, so the match is the
    maximal non-whitespace run from here — provided at least one
    character follows the scheme. -/
def matchUrl (s : List Char) : Option Nat :=
  let schemeLen :=
    if s.take 8 == "https://".toList then some 8
    else if s.take 7 == "http://".toList then some 7
    else none
  match schemeLen with
  | none => none
  | some sl =>
    let r := (s.takeWhile (fun c => !isWs c)).length
    if r > sl then some r else none

/-- `[[...]]` / `［［...］］` markers inside prose. -/
def matchMarkerAt (s : List Char) (o cl : Char) : Option Nat :=
  if s.take 2 == [o, o] then
    let body := s.drop 2
    let blen := nonRunLen body cl
    if body[blen]? == some cl && body[blen + 1]? == some cl then
      some (2 + blen + 2)
    else none
  else none

/-- Raw HTML tags: `<[^>]*>`. -/
def matchTag (s : List Char) : Option Nat :=
  if s.head? == some '<' then
    let blen := nonRunLen (s.drop 1) '>'
    if (s.drop 1)[blen]? == some '>' then some (1 + blen + 1)
    else none
  else none

/-- Try the alternatives of INLINE_PROTECTED_RE_ at the head of
    `s`, in the regex's own order. Returns the match length and
    the segments the match produces. -/
def tryMatchAt (s : List Char) : Option (Nat × List (List Char × Bool)) :=
  match matchCodeSpan s with
  | some len => some (len, [(s.take len, false)])
  | none =>
  match matchLink s with
  | some r => some r
  | none =>
  match matchUrl s with
  | some len => some (len, [(s.take len, false)])
  | none =>
  match matchMarkerAt s '[' ']' with
  | some len => some (len, [(s.take len, false)])
  | none =>
  match matchMarkerAt s '［' '］' with
  | some len => some (len, [(s.take len, false)])
  | none =>
  match matchTag s with
  | some len => some (len, [(s.take len, false)])
  | none => none

/-- Find the leftmost inline match: (prose length before it,
    match length, its pieces). -/
def findMatch (s : List Char) : Option (Nat × Nat × List (List Char × Bool)) :=
  let rec go (pos : Nat) : List Char → Option (Nat × Nat × List (List Char × Bool))
    | [] => none
    | rest@(c :: t) =>
      match tryMatchAt rest with
      | some (len, pieces) => some (pos, len, pieces)
      | none => go (pos + 1) t
  go 0 s

def scanInlineGo : Nat → List Char → List (List Char × Bool)
  | 0, s => if s.isEmpty then [] else [(s, true)]
  | fuel + 1, s =>
    match findMatch s with
    | none => if s.isEmpty then [] else [(s, true)]
    | some (pre, len, pieces) =>
      (if pre > 0 then [(s.take pre, true)] else []) ++ pieces ++
      scanInlineGo fuel (s.drop (pre + len))

def scanInline (s : List Char) : List (List Char × Bool) :=
  scanInlineGo (s.length + 1) s

/-- Split one prose line: structural prefix first (protected),
    then inline pieces. -/
def splitProseLine (line : List Char) : List (List Char × Bool) :=
  let p := matchPrefix line
  (if p > 0 then [(line.take p, false)] else []) ++ scanInline (line.drop p)

/-! ## Block assembly: splitMarkdownSegments_ -/

def lineSegs (l : List Char) : List (List Char × Bool) :=
  if lineProtected l then [(l, false)] else splitProseLine l

def joinLines : List (List Char) → List (List Char × Bool)
  | [] => []
  | [l] => lineSegs l
  | l :: ls => lineSegs l ++ [("\n".toList, false)] ++ joinLines ls

/-- Segments for one text block, plus the fence state after it. -/
def segmentsOfBlock (part : List Char) (fenceOpen : Bool) :
    List (List Char × Bool) × Bool :=
  let lines := splitLines part
  let (touched, newOpen) := scanFence lines fenceOpen
  if touched then ([(part, false)], newOpen)
  else if blockProtected part then ([(part, false)], newOpen)
  else (joinLines lines, newOpen)

/-- `splitMarkdownSegments_(markdown)`: segments whose
    concatenation is the input, byte for byte; `prose = true`
    marks the only pieces that may enter a translation call. -/
def splitSegments (md : List Char) : List (List Char × Bool) :=
  let rec go : List (List Char × Bool) → Bool → List (List Char × Bool)
    | [], _ => []
    | (txt, isSep) :: rest, isOpen =>
      if txt.isEmpty then go rest isOpen
      else if isSep then (txt, false) :: go rest isOpen
      else
        let (segs, newOpen) := segmentsOfBlock txt isOpen
        segs ++ go rest newOpen
  go (tokenize md) false

/-! ## Batch planning: the LanguageApp calls of translateMarkdown_ -/

/-- Approximation of the JS letter test `/\p{L}/u`, exact for the
    character ranges the differential harness generates: ASCII
    letters, hiragana, katakana (excluding ・ U+30FB, which is
    punctuation), CJK ideographs, full-width letters and 々. -/
def isLetterC (c : Char) : Bool :=
  let n := c.toNat
  (0x41 ≤ n && n ≤ 0x5A) || (0x61 ≤ n && n ≤ 0x7A) ||
  (0x3041 ≤ n && n ≤ 0x3096) || (0x30A1 ≤ n && n ≤ 0x30FA) ||
  n == 0x30FC || (0x30FD ≤ n && n ≤ 0x30FF) ||
  (0x4E00 ≤ n && n ≤ 0x9FFF) ||
  (0xFF21 ≤ n && n ≤ 0xFF3A) || (0xFF41 ≤ n && n ≤ 0xFF5A) ||
  n == 0x3005

def hasLetter (s : List Char) : Bool := s.any isLetterC

/-- The batching loop of `translateMarkdown_`, over the
    letter-bearing units in order. Note the faithful quirk: when
    a batch is flushed, the new batch's length counter starts at
    the flushed unit's `need` — which was computed with the
    newline +1 for the previous batch — so post-flush batches
    effectively budget 3999, not 4000. -/
def batchGo : List (List Char) → List (List Char) → Nat →
    List (List (List Char)) → List (List (List Char))
  | [], batch, _, acc => (if batch.isEmpty then acc else batch :: acc).reverse
  | u :: rest, batch, len, acc =>
    let need := u.length + (if batch.isEmpty then 0 else 1)
    if !batch.isEmpty && len + need > 4000 then
      batchGo rest [u] need (batch :: acc)
    else
      batchGo rest (batch ++ [u]) (len + need) acc

/-- The exact strings that would be passed to LanguageApp, in
    order, for this document: prose segments split at 4000,
    letterless fragments skipped, the rest batched newline-joined
    under the 4000 budget. -/
def planCalls (md : List Char) : List String :=
  let units := (splitSegments md).filterMap
    (fun seg => if seg.2 then some seg.1 else none)
    |>.flatMap (splitLongText 4000)
  let letterUnits := units.filter hasLetter
  (batchGo letterUnits [] 0 []).map
    (fun batch => String.intercalate "\n" (batch.map List.asString))

/-! ## String-level boundary used by the oracle -/

def segments (md : String) : List (String × Bool) :=
  (splitSegments md.toList).map (fun seg => (List.asString seg.1, seg.2))

def splitLong (text : String) (limit : Nat) : List String :=
  (splitLongText limit text.toList).map List.asString

def calls (md : String) : List String := planCalls md.toList

end BlogGasSpec
