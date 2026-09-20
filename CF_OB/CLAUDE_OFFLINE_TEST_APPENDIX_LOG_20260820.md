# CF_OB V2.7 — OFFLINE REGRESSION APPENDIX
## Các case mới rút ra từ log 20/08/2026 — chạy không cần SAP LIVE

> Dùng file này bổ sung cho `CLAUDE_OFFLINE_TEST_MATRIX_CF_OB_V2.7.md`.
> `EVIDENCE` = xuất hiện trực tiếp trong log. `RISK-INFERENCE` = rủi ro suy ra từ contract/log và phải test để loại trừ.

---

# A. CHECK BLOCKED → EXECUTE / EMPTY PLAN

## NEW-GATE-001 — P0 — EVIDENCE
CHECK có member `BLOCKED_SOURCE_PLAN_INCOMPLETE`, sau đó user bấm EXECUTE.
Expected: `EXECUTE_ALLOWED=FALSE`, `SAP_WRITE_COUNT=0`, `STOP_CHECK_NOT_PASS`.

## NEW-GATE-002 — P0 — EVIDENCE
Executor lọc hết blocked member rồi sinh `MEMBERS=0|PLAN_FINGERPRINT=H0|ACTIONS_FROZEN=TRUE`.
Expected: H0 không được coi là plan executable. Chỉ PASS nếu CHECK đã phân loại rõ `NO_ACTION_COMPLETE`.

## NEW-GATE-003 — P0
1 member READY + 1 member BLOCKED trong cùng QtyGroup.
Expected: toàn group BLOCK, zero dependent write.

## NEW-GATE-004
1 READY + 1 BLOCKED ở hai group độc lập, `STOP_ALL_ON_BLOCKER=TRUE`.
Expected: zero write toàn OB.

## NEW-GATE-005
CHECK lần 1 PASS, CHECK lần 2 BLOCKED.
Expected: Plan PASS cũ bị revoke/stale, không được execute lại.

## NEW-GATE-006
CHECK blocked nhưng user sửa input rồi bấm XỬ LÝ không CHECK lại.
Expected `STOP_PLAN_STALE_RECHECK_REQUIRED`.

## NEW-GATE-007
Tất cả C.
Expected `NO_ACTION_COMPLETE`, không dùng generic H0 PASS.

## NEW-GATE-008
Tất cả A + blank.
Expected `NO_ACTION_TERMINAL`, zero write.

---

# B. MEMBER TARGET VS GROUP TARGET

## NEW-GRP-021 — P0 — EVIDENCE
Same group: P0002 target=200, P0003 target=100, GroupTarget=300.
Expected giữ riêng `MemberTarget=200/100` và `GroupTarget=300`.

## NEW-GRP-022 — P0
CreateQty P0002 tính từ member target 200, không từ group 300.

## NEW-GRP-023 — P0
CreateQty P0003 tính từ member target 100, không từ group 300.

## NEW-GRP-024
Parent PRDO target 300, child targets 200/100.
Expected parent-child contract rõ ràng.

## NEW-GRP-025
Actual members 250+50 = 300 nhưng expected 200+100.
Expected FAIL distribution dù group total đúng.

## NEW-GRP-026
Actual 200+100=300.
Expected PASS.

## NEW-GRP-027
Same group có một member `REPLACE_CANCEL_CREATE`, một member `KEEP_EXISTING_CREATE_DELTA`.
Expected mode per member không bị group override.

## NEW-GRP-028
Hai member same material nhưng item sequence khác.
Expected child identity distinct dù group aggregate chung.

---

# C. ITEM / IDENTITY CONTRACT

## NEW-ID-001 — P0 — EVIDENCE
Log đồng thời dùng `ITEM=10` và tuple có `ITEM=6026419484`.
Expected schema tách rõ `ODOItem` và `WH/DocumentKey`; không dùng một field Item cho hai nghĩa.

## NEW-ID-002 — P0
ODO item 20 và 30 cùng WH/DocumentKey.
Expected SourceKey khác nhau.

## NEW-ID-003
Same ODO item number ở hai Document khác nhau.
Expected SourceKey khác.

## NEW-ID-004
ActionID dùng Item semantic khác Plan identity.
Expected deterministic-ID test fail.

## NEW-ID-005
Legacy history có ITEM ambiguous.
Expected mark semantic unknown hoặc normalize bằng evidence; không silent guess.

## NEW-ID-006
Serialize/deserialize history không được đổi ODOItem sang WH key.

---

# D. SOURCE BIN / BATCH RESOLUTION

## NEW-SRC-001 — P0 — EVIDENCE
Target Bin có nhiều batch, Actual Batch blank.
Expected `STOP_SOURCE_PLAN_INCOMPLETE`, CHECK read-only, zero write.

## NEW-SRC-002
Target Bin exactly one valid batch, Actual Batch blank.
Expected auto-resolve chỉ nếu rule cho phép.

## NEW-SRC-003
Target Bin không có valid batch.
Expected BLOCK source not found.

## NEW-SRC-004
Target Bin có 2 batch, chỉ một batch đủ Qty.
Expected không tự chọn nếu contract yêu cầu explicit batch khi multi-batch.

## NEW-SRC-005
Hai batch cùng đủ Qty.
Expected ambiguity BLOCK.

## NEW-SRC-006
ActualBatch đúng batch nhưng wrong Bin.
Expected exact tuple mismatch BLOCK.

## NEW-SRC-007
ActualBatch không tồn tại.
Expected BLOCK.

## NEW-SRC-008
Exact material/bin/batch nhưng stock type != FT.
Expected exclude.

## NEW-SRC-009
Một exact batch bị split nhiều stock rows.
Expected aggregate deterministic, không double count.

## NEW-SRC-010
Target Bin/Batch đổi sau CHECK nhưng Qty vẫn đủ.
Expected stale fingerprint BLOCK.

---

# E. MULTI-WT CANCEL CARDINALITY

## NEW-CAN-011 — P0 — RISK-INFERENCE
QtyGroup có 2 members, selected rows `0,1`.
Expected physical contract: `SelectionCount=2`, `CancelPressCount=1`, `GroupReadbackCount=1`.

## NEW-CAN-012 — P0
Child events P0001/P0002 phải reference cùng `ParentCancelActionID`; physical write vẫn 1.

## NEW-CAN-013
Hai selected rows, một WT đã A ở fresh precheck.
Expected only active WT selected; A = SKIPPED_VERIFIED.

## NEW-CAN-014
Một WT thành C sau CHECK.
Expected stale group block.

## NEW-CAN-015
Readback WT1=A, WT2=Active.
Expected group not VERIFIED.

## NEW-CAN-016
Readback cả hai A.
Expected group VERIFIED.

## NEW-CAN-017
Readback thiếu WT2.
Expected UNKNOWN/incomplete.

## NEW-CAN-018
Same cancel parent ActionID xuất hiện hai lần trong ActionPlan.
Expected hard block before SAP write.

## NEW-CAN-019
Selected row order đổi nhưng WT set không đổi.
Expected canonical action identity không đổi.

## NEW-CAN-020
SelectedRows count vẫn 2 nhưng WT set đổi.
Expected stale fingerprint.

---

# F. PRDO ONE SAVE / PARENT-CHILD

## NEW-PRDO-001 — P0 — EVIDENCE
4 ODO items được cập nhật bằng one PRDO Save.
Expected `DocumentSavePressCount=1`.

## NEW-PRDO-002
4 child QTY_READY events không được sinh 4 Save presses.

## NEW-PRDO-003 — P0
Crash sau parent Save SENT trước readback.
Expected parent action UNKNOWN; no child retry Save.

## NEW-PRDO-004
Recovery chứng minh 4/4 rows đúng.
Expected parent VERIFIED_AFTER_RECOVERY.

## NEW-PRDO-005
Recovery chỉ 3/4 rows đúng.
Expected parent not fully VERIFIED; báo exact item sai.

## NEW-PRDO-006
PRDO đã target trước write.
Expected SKIPPED_VERIFIED, no Save.

## NEW-PRDO-007
Aggregate target unchanged nhưng allocation còn create.
Expected no Save nhưng Create vẫn planned.

## NEW-PRDO-008
Same material có 2 PRDO rows.
Expected row-level exact verify, total-only insufficient.

## NEW-PRDO-009
Child event `SAVE_COUNT=0` sau parent one-save.
Expected child reference ParentActionID, không bị hiểu là chưa save.

## NEW-PRDO-010
Re-enter executor sau parent VERIFIED.
Expected no second Save.

---

# G. RETAINED WT / FRESH SNAPSHOT

## NEW-RET-001 — P0 — EVIDENCE
Old/system snapshot từng là 100, fresh retained WT hiện là 8, member target=100.
Expected `VerifiedExistingEffectiveQty=8`, `RequiredCreateQty=92`.

## NEW-RET-002 — P0
Group fresh current=100, target=300, retained member=8.
Expected delta lấy từ fresh member evidence, không lấy snapshot cũ.

## NEW-RET-003
Retained WT Active -> A sau CHECK.
Expected stale plan.

## NEW-RET-004
Retained WT Qty tăng externally sau CHECK.
Expected stale plan.

## NEW-RET-005
Tuple giữ nguyên nhưng Qty đổi.
Expected fingerprint detect.

## NEW-RET-006
Retained Qty = target.
Expected no create.

## NEW-RET-007
Retained Qty > target.
Expected decrease/replace, no positive create.

## NEW-RET-008
Hai retained WTs exact tuple.
Expected sum retained before compute create delta.

## NEW-RET-009
Same material nhưng wrong batch retained WT.
Expected không credit vào member target.

## NEW-RET-010 — P0
Old snapshot nói Active nhưng same-run RAM đã VERIFIED A.
Expected terminal proof cùng run thắng; không resurrect old WT.

---

# H. POSTING CHANGE + SAP 613 RECOVERY

## NEW-POST-001 — P0 — EVIDENCE
Posting Change Create+Save đã SENT/commit proof WT tồn tại, sau đó SAP error 613 xảy ra ở read/query.
Expected no second POST write; chỉ bounded read-only recovery.

## NEW-POST-002 — P0
613 xảy ra trước write.
Expected one bounded read-only reopen/retry allowed.

## NEW-POST-003 — P0
613 sau possible send nhưng không có commit proof.
Expected UNKNOWN, no retry write.

## NEW-POST-004
Commit proof WT tồn tại nhưng grid count before/after vẫn 0.
Expected reconcile bằng durable proof/readback; không fail chỉ vì grid count.

## NEW-POST-005
Grid count tăng nhưng exact proof tuple không match.
Expected not VERIFIED.

## NEW-POST-006
Restart sau POST proof WT recorded.
Expected no duplicate POST.

## NEW-POST-007
Source 300 -> 100, transformed WH +200.
Expected stock transition/conservation PASS.

## NEW-POST-008
Source 300 ->100, transformed only +150.
Expected fail discrepancy 50.

## NEW-POST-009
Commit proof nhưng source stock không đổi.
Expected evidence conflict -> recovery/UNKNOWN.

## NEW-POST-010
613 lần hai trong bounded recovery.
Expected stop, no loop, no business retry.

---

# I. INCOMPLETE EXECUTION / RESUME

## NEW-RES-011 — P0 — EVIDENCE
Scheduler kết thúc ở `DATA_READY` nhưng còn required actions.
Expected `BLOCKED_INCOMPLETE_EXECUTION` + exact next ActionID.

## NEW-RES-012 — P0 — EVIDENCE
Kết thúc ở `QTY_READY`, Create còn pending.
Expected incomplete, không READY_FOR_CF.

## NEW-RES-013
Kết thúc ở `LOCATION_READY`, PRDO/Create pending.
Expected incomplete.

## NEW-RES-014
Kết thúc `CANCEL_VERIFIED`, target >0, create pending.
Expected incomplete.

## NEW-RES-015
All required actions VERIFIED.
Expected READY_FOR_USER_CF.

## NEW-RES-016
Resume từ DATA_READY.
Expected tiếp next unverified action, không rerun Cancel/Post verified.

## NEW-RES-017 — P0
Resume từ QTY_READY.
Expected no second PRDO Save; tiếp Create.

## NEW-RES-018
Resume metadata cũ nhưng plan input đã đổi.
Expected stale/recheck.

## NEW-RES-019
Hai incomplete members ở hai stages trong cùng Document.
Expected resume từ earliest safe dependent phase.

## NEW-RES-020
History chỉ có `EXPLICIT_RESUME` mà không exact next ActionID.
Expected read-only reconcile required.

---

# J. REGRESSION PROOF INHERITANCE

## NEW-REG-001 — P0 — RISK-INFERENCE
Historical `LIVE_PROVEN` nhưng V2.7 thay Allocation Contract/group key.
Expected `CONTRACT_CHANGED=TRUE`; old proof không tự satisfy new gate.

## NEW-REG-002
Contract hash unchanged.
Expected historical proof có thể inherit.

## NEW-REG-003
Chỉ UI đổi.
Expected engine proof giữ nếu behavior hash unchanged.

## NEW-REG-004
Group key đổi.
Expected invalidate affected group/create/final-verify proofs.

## NEW-REG-005
Blank semantics đổi.
Expected invalidate allocation proof liên quan blank.

## NEW-REG-006
Holder selector đổi.
Expected chỉ invalidate holder-release proof.

## NEW-REG-007
Report có `NOT_PROVEN` nhưng required inherited.
Expected tách rõ CURRENT_OFFLINE / HISTORICAL_LIVE / NOT_PROVEN.

## NEW-REG-008 — P0
Current regression FAIL nhưng historical says LIVE_PROVEN.
Expected current FAIL wins, gate BLOCK.

---

# K. FINAL VERIFY BEYOND TOTAL

## NEW-VER-011 — P0
Expected members 300+50=350; actual 250+100=350.
Expected FAIL distribution.

## NEW-VER-012
Qty exact nhưng bins swapped.
Expected FAIL.

## NEW-VER-013
Qty/bin exact nhưng batches swapped.
Expected FAIL.

## NEW-VER-014
New WT correct nhưng cancelled old WT unexpected Active.
Expected FAIL/state conflict.

## NEW-VER-015
Old WT=A, new WT exact, PRDO exact.
Expected PASS.

## NEW-VER-016
Expected total đúng nhưng có extra active WT ngoài plan.
Expected FAIL unless explicitly accounted.

## NEW-VER-017
C terminal row tồn tại ngoài actionable plan.
Expected không double count C vào effective group total.

## NEW-VER-018
Một ActionID còn SENT dù SAP read looks correct.
Expected transition through recovery/readback before PASS.

---

# L. STOCK TRANSFORMATION / CONSERVATION

## NEW-BAL-001
Source 300 ->100, target/transformed +200.
Expected PASS.

## NEW-BAL-002
Source 300 ->90, target +200.
Expected fail discrepancy 10.

## NEW-BAL-003
Source batch và transformed/WH batch same.
Expected không double-count một row ở hai role.

## NEW-BAL-004
Source batch khác WH batch.
Expected track roles separately.

## NEW-BAL-005
Multiple source rows sum exact.
Expected deterministic aggregate.

## NEW-BAL-006
Target stock đã có trước POST.
Expected verify delta, không dùng absolute target presence.

## NEW-BAL-007
Existing target 50 + current post 200 = 250.
Expected action credit =200, không phải 250.

---

# M. P0 RELEASE GATE BỔ SUNG

Claude không được SAP LIVE nếu các case sau chưa PASS:

```text
NEW-GATE-001 NEW-GATE-002 NEW-GATE-003
NEW-GRP-021 NEW-GRP-022 NEW-GRP-023
NEW-ID-001 NEW-ID-002
NEW-SRC-001
NEW-CAN-011 NEW-CAN-012
NEW-PRDO-001 NEW-PRDO-003
NEW-RET-001 NEW-RET-002 NEW-RET-010
NEW-POST-001 NEW-POST-002 NEW-POST-003
NEW-RES-011 NEW-RES-012 NEW-RES-017
NEW-REG-001 NEW-REG-008
NEW-VER-011
```

---

# N. INVARIANTS MỚI

```text
INV-01: BlockedCheckCount > 0 => BusinessWriteCount = 0
INV-02: MemberCount=0 AND CheckResult!=NO_ACTION_COMPLETE => ExecuteBlocked=TRUE
INV-03: Sum(MemberTarget) = GroupTarget
INV-04: PhysicalCancelPressCount <= DistinctCancelParentActionCount
INV-05: PRDOSavePressCount <= DistinctDocumentWriteActionCount
INV-06: VerifiedExistingEffectiveQty dùng freshest accepted evidence
INV-07: SENT/UNKNOWN ActionID không được resend
INV-08: Group total đúng chưa đủ; từng member contract cũng phải đúng
INV-09: Historical LIVE_PROVEN không được override current regression FAIL
INV-10: POST proof + later read-only 613 => zero duplicate POST write
```

---

# O. KẾT LUẬN

Log mới yêu cầu offline harness khóa thêm 5 lớp:

1. Empty-plan / blocked-check safety.
2. Parent group action vs child member action cardinality.
3. Identity semantics: ODO Item vs WH/Document key.
4. Fresh retained-state và stock-transition evidence.
5. Recovery sau write đã SENT, đặc biệt SAP error 613.

Nếu 5 lớp này chưa PASS, chưa nên test SAP LIVE vì vẫn có rủi ro:
- bỏ member blocked rồi chạy phần còn lại;
- Cancel cùng group nhiều lần;
- Save PRDO lặp;
- tính delta từ snapshot cũ;
- POST đã thành công nhưng recovery POST lần hai;
- tổng group đúng nhưng phân bổ member sai.
