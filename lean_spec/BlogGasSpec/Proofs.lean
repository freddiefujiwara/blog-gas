/-
  Theorems and #guard checks for the executable specification.
  `lake build` elaborates this file, so the proofs and the guards
  are the test suite, as in the asset prototype.
-/
import BlogGasSpec.Spec

namespace BlogGasSpec

/-! ## lastBoundary stays inside the window -/

theorem lastBoundaryGo_spec : ∀ (i : Nat) (acc : Option Nat)
    (w : List Char),
    (∀ j, acc = some j → j < i) →
    ∀ j, lastBoundaryGo i acc w = some j → j < i + w.length := by
  intro i acc w
  induction w generalizing i acc with
  | nil =>
    intro hacc j h
    simp [lastBoundaryGo] at h
    subst h
    have := hacc j rfl
    omega
  | cons c t ih =>
    intro hacc j h
    simp only [lastBoundaryGo] at h
    have hacc' : ∀ x, (if (c == '。' || c == '！' || c == '？' ||
        c == '、' || (c == '.' && t.head? == some ' ')) then some i
        else acc) = some x → x < i + 1 := by
      intro x hx
      by_cases hb : (c == '。' || c == '！' || c == '？' || c == '、' ||
          (c == '.' && t.head? == some ' '))
      · simp [hb] at hx
        subst hx
        omega
      · simp [hb] at hx
        have := hacc x hx
        omega
    have := ih (i + 1) _ hacc' j h
    simp at this ⊢
    omega

theorem lastBoundary_lt {w : List Char} {j : Nat}
    (h : lastBoundary w = some j) : j < w.length := by
  have := lastBoundaryGo_spec 0 none w (fun j hj => by simp at hj) j h
  simpa using this

/-! ## chooseCut bounds -/

theorem chooseCut_pos {limit : Nat} (h : 1 ≤ limit) (w : List Char) :
    1 ≤ chooseCut limit w := by
  unfold chooseCut
  cases hb : lastBoundary w with
  | none => exact h
  | some i =>
    by_cases hc : i >= (limit * 3) / 10
    · simp [hc]
    · simp [hc, h]

theorem chooseCut_le {limit : Nat} (w : List Char) :
    chooseCut limit w ≤ max limit w.length := by
  unfold chooseCut
  cases hb : lastBoundary w with
  | none => exact Nat.le_max_left _ _
  | some i =>
    have hi := lastBoundary_lt hb
    by_cases hc : i >= (limit * 3) / 10
    · simp only [hc, ↓reduceIte]
      omega
    · simp only [hc, ↓reduceIte]
      exact Nat.le_max_left _ _

/-! ## splitLongText: conservation and the bound -/

theorem splitLongGo_concat : ∀ (limit fuel : Nat) (rest : List Char),
    rest.length ≤ fuel → (splitLongGo limit fuel rest).flatten = rest := by
  intro limit fuel
  induction fuel with
  | zero =>
    intro rest hlen
    cases rest with
    | nil => rfl
    | cons c t => simp at hlen
  | succ f ih =>
    intro rest hlen
    cases rest with
    | nil => rfl
    | cons c t =>
      by_cases hlim : limit < 1
      · simp [splitLongGo, hlim]
      · by_cases hle : (c :: t).length ≤ limit
        · simp only [splitLongGo, hlim, hle, ↓reduceIte,
            List.flatten_cons, List.flatten_nil, List.append_nil]
        · have hcut1 : 1 ≤ chooseCut limit ((c :: t).take limit) :=
            chooseCut_pos (by omega) _
          have hcut2 : chooseCut limit ((c :: t).take limit) ≤ limit := by
            have h := chooseCut_le (limit := limit) ((c :: t).take limit)
            have hlen' : ((c :: t).take limit).length = limit := by
              rw [List.length_take]
              omega
            omega
          have hdrop :
              ((c :: t).drop (chooseCut limit ((c :: t).take limit))).length ≤ f := by
            simp only [List.length_drop, List.length_cons] at *
            omega
          simp only [splitLongGo, hlim, hle, ↓reduceIte, List.flatten_cons]
          rw [ih _ hdrop, List.take_append_drop]

theorem splitLong_concat (limit : Nat) (text : List Char) :
    (splitLongText limit text).flatten = text :=
  splitLongGo_concat limit text.length text (by omega)

theorem splitLongGo_bound : ∀ (limit fuel : Nat) (rest : List Char),
    rest.length ≤ fuel → 1 ≤ limit →
    ∀ p ∈ splitLongGo limit fuel rest, p.length ≤ limit ∧ p ≠ [] := by
  intro limit fuel
  induction fuel with
  | zero =>
    intro rest hlen _ p hp
    cases rest with
    | nil =>
      rw [show splitLongGo limit 0 [] = [] from rfl] at hp
      cases hp
    | cons c t => simp at hlen
  | succ f ih =>
    intro rest hlen hlim p hp
    cases rest with
    | nil =>
      rw [show splitLongGo limit (f + 1) [] = [] from rfl] at hp
      cases hp
    | cons c t =>
      have hnl : ¬ limit < 1 := by omega
      by_cases hle : (c :: t).length ≤ limit
      · simp only [splitLongGo, hnl, hle, ↓reduceIte, List.mem_singleton] at hp
        subst hp
        exact ⟨hle, by simp⟩
      · have hcut1 : 1 ≤ chooseCut limit ((c :: t).take limit) :=
          chooseCut_pos hlim _
        have hcut2 : chooseCut limit ((c :: t).take limit) ≤ limit := by
          have h := chooseCut_le (limit := limit) ((c :: t).take limit)
          have hlen' : ((c :: t).take limit).length = limit := by
            rw [List.length_take]
            omega
          omega
        have hdrop :
            ((c :: t).drop (chooseCut limit ((c :: t).take limit))).length ≤ f := by
          simp only [List.length_drop, List.length_cons] at *
          omega
        simp only [splitLongGo, hnl, hle, ↓reduceIte, List.mem_cons] at hp
        cases hp with
        | inl h =>
          subst h
          refine ⟨?_, ?_⟩
          · simp only [List.length_take]
            omega
          · intro hcon
            have hlen0 : ((c :: t).take
                (chooseCut limit ((c :: t).take limit))).length = 0 :=
              List.length_eq_zero_iff.mpr hcon
            simp only [List.length_take] at hlen0
            omega
        | inr h => exact ih _ hdrop hlim p h

theorem splitLong_bound (limit : Nat) (hlim : 1 ≤ limit) (text : List Char) :
    ∀ p ∈ splitLongText limit text, p.length ≤ limit ∧ p ≠ [] :=
  splitLongGo_bound limit text.length text (by omega) hlim

/-! ## #guard checks: segmentation and batching corner cases -/

-- A fenced block is one protected segment; prose around it splits.
#guard (segments "本文です。\n\n```js\nconst x = 1;\n```\n\n次の段落。") ==
  [("本文です。", true), ("\n\n", false),
   ("```js\nconst x = 1;\n```", false), ("\n\n", false),
   ("次の段落。", true)]

-- A bare-URL block is protected whole.
#guard (segments "見て。\n\nhttps://example.com/a?b=c\n\n以上。") ==
  [("見て。", true), ("\n\n", false),
   ("https://example.com/a?b=c", false), ("\n\n", false),
   ("以上。", true)]

-- Link text is prose; brackets and destination are protected.
#guard (segments "[川口の店](https://example.com/shop)に行った。") ==
  [("[", false), ("川口の店", true),
   ("](https://example.com/shop)", false), ("に行った。", true)]

-- An unclosed fence protects the rest of the document.
#guard (segments "前。\n\n```\nコード\n本文も含む") ==
  [("前。", true), ("\n\n", false), ("```\nコード\n本文も含む", false)]

-- The batching quirk: after a flush the length counter keeps the
-- flushed unit's +1, so [3000, 1500, 2499] lands in three batches
-- where an exact accounting would give two.
#guard (batchGo
    [List.replicate 3000 'あ', List.replicate 1500 'あ',
     List.replicate 2499 'あ'] [] 0 []).length == 3

-- splitLongText cuts at the last sentence boundary inside the
-- window when it is at or beyond 30% of the limit.
#guard splitLong "あ。いい。うううううううう" 10 ==
  ["あ。いい。", "うううううううう"]

end BlogGasSpec
