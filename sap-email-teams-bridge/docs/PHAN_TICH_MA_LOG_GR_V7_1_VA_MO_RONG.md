# Phân tích MA Log GR Tool V7.1 và đề xuất mở rộng bridge

Ngày: 24/10/2026. Phạm vi: đọc workbook và phân tích VBA tĩnh, đối chiếu source bridge; chưa chạy macro, truy vấn hoặc ghi SAP, chưa thêm handler mới.

## 1. Nguồn và kết luận

Nguồn chính: `C:\Users\taikt\Downloads\240101_MA Log_GR Tool_V7_1.xlsm`, 9.463.765 bytes, SHA-256 `87201810639d35760d2d9d0dcab9bfdb9ec04690e5ad0bf25a4133accbad8bb7`.

Hash này trùng workbook được ghi trong `G:\SAP\MA_LOG_GR_V7_1_4_README.md`, đi kèm archive `240101_MA_Log_GR_Tool_V7_1_4_LZMA.zip`. Đây là bản workbook nén, không phải một tool Python riêng. Các file Module*.bas khác trong G:\SAP không được coi là source workbook này khi chưa chứng minh khớp.

Workbook có hai sheet `GR_EWM_V1` và `Template`. VBA gồm `modSAP_Functions` (32 Sub), `modSupports` và các module workbook/sheet. Số dòng VBA bên dưới theo source trích xuất đã chuẩn hóa xuống dòng; không phải dòng Excel.

Đây là công cụ thao tác GR/inbound/PO và dữ liệu kho, có nhiều chức năng ghi thật. Nên bổ sung trước sáu chức năng tra cứu: tìm chứng từ GR, kiểm tra inbound, tìm inbound, kiểm tra PO/SA, xem queue lỗi và xem shipment. Sau đó mới ghép thành báo cáo kiểm tra trước GR. Báo cáo ZEWMR137 còn thiếu định nghĩa cột/nghiệp vụ để triển khai.

Workbook này không phải master vật tư thay thế `data/DATA_MASTER.xlsx`. Không mang toàn workbook hoặc Excel COM vào worker chỉ để tái sử dụng VBA.

## 2. Chức năng có thể tái sử dụng ý tưởng

| Macro gốc | Nguồn SAP | Dữ liệu thực sự đọc | Khác biệt so với bridge hiện tại |
|---|---|---|---|
| `Get_Mat_Doc`, 679–814 | MB51, lọc reference XBLNR | `XBLNR`, `MBLNR`, `MATNR` | Bridge có MB51 theo material/ngày, chưa tìm chứng từ theo inbound/reference |
| `Check_Inbound_Qty`, 815–976 | VL06I, item view | `VBELN`, `MATNR`, `LFIMG`, `LGORT`, `WBSTA`, `S_VGBEL`, `S_VGPOS` | Cần tra theo số IB và đối chiếu quantity; IB_TREO hiện chỉ là danh sách chưa GR |
| `Find_Inbound_VL06I`, 977–1166 | VL06I | Thêm `LIFEX`, PO/SA, item, số lượng và trạng thái | Tìm IB theo PO/SA/material hoặc external reference; không chỉ tìm theo material |
| `Check_Queued`, 1167–1235 | /SCWM/MON, node MQUEUE0001 | `BUSOBJ_KEY`, `ERRMESS`; `QSTATE` đang comment trong bản gốc | Monitor history-bin hiện tại không phải màn hình queue lỗi |
| `Find_Inbound_ZMME240`, 1864–2016 | ZMME240 | `TRXID1`, `PART_NAME`, `QUANTITY`, `HOUSEBILL`, `VGBEL`, `STATUS` | Báo cáo tùy biến cho shipment/IB chưa có trong bridge |
| `Check_PO_SA_Availability`, 2119–2424 | ME2N | `EBELN`, `EBELP`, `EMATN`, `MGLIEF`, `FRGKE`, `LOEKZ` | ME2N hiện tại tập trung STO theo material; chưa kiểm PO/SA và quantity yêu cầu |
| `Report_ZEWMR137`, 2654–2678 | ZEWMR137 | Chỉ mở báo cáo, execute và export XLSX | Chưa có định nghĩa cột, filter và KPI; chưa đủ bằng chứng tạo báo cáo mới |

Tên kỹ thuật và layout cần kiểm tra trên SAP đích. Đặc biệt `MGLIEF`, mã release và shipment status không được tự gán nghĩa dựa vào tên biến hoặc suy từ STO.

## 3. Thiết kế sáu lệnh mới

Tên lệnh là đề xuất, chưa tồn tại trong bridge. Cú pháp `<...>` dưới đây là placeholder.

### A. TIM_GR_DOC — tìm chứng từ GR

Ví dụ: `SAP tìm chứng từ GR inbound <IB> mã <material> từ <ngày> đến <ngày>`.

- Nhận inbound/reference nguyên bản, material, plant và phạm vi ngày; hỗ trợ nhiều IB có giới hạn.
- Mở rộng MB51 hiện có, trả tất cả chứng từ khớp cùng item, năm chứng từ, posting date, movement, quantity và đơn vị nếu các cột nguồn cung cấp.
- Mã chứng từ phải đi kèm năm/item khi dùng làm khóa; phân biệt posting date và document date.
- Nếu dùng reference suy ra từ IB, trả rõ reference gốc và quy tắc chuyển đổi đã xác minh. Không mặc định cắt 9 ký tự hoặc thêm 0 vào mọi loại chứng từ.
- “Tìm thấy chứng từ tham chiếu” không tự chứng minh GR còn hiệu lực. Chỉ kết luận nhận ròng sau khi đọc đủ movement/reversal/return theo quy tắc nghiệp vụ đã xác nhận.

### B. CHECK_IB — kiểm tra inbound theo số

Ví dụ: `SAP kiểm tra inbound <IB> mã <material> số lượng <qty> <unit>`.

- Trả IB/item, material, quantity, unit, SLoc, trạng thái nguồn, PO/SA và item tham chiếu.
- Dùng IB/item/material làm định danh; quantity là giá trị đối chiếu. Nếu chưa có item và có nhiều kết quả, giữ đầy đủ candidates thay vì chọn dòng đầu.
- Hiển thị `expectedQty`, `actualQty`, `delta` khi đủ dữ liệu và cùng đơn vị. Thiếu quantity yêu cầu vẫn cho phép tra trạng thái, không tạo chênh lệch giả.
- Đọc cả trạng thái chưa GR, một phần, hoàn tất theo nguồn. Không dùng handler chỉ lọc A để trả lời câu hỏi trạng thái toàn bộ IB.
- Lời trả “SL khớp dữ liệu inbound” không được đồng nghĩa “được phép post GR”; còn phụ thuộc PO, batch, quyền và các điều kiện khác.

### C. TIM_IB — tìm inbound chưa biết số

Ví dụ: `SAP tìm inbound PO <PO> mã <material>` hoặc `SAP tìm inbound bill <reference> mã <material>`.

- Bắt đầu bằng VL06I, dùng material + PO/SA/item hoặc external reference. Quantity là tiêu chí đối chiếu bổ sung.
- Trả số IB/item, PO/item, external reference, quantity/unit, SLoc và trạng thái; có cờ nhiều kết quả.
- Cho phép giao tách nhiều đợt/nhiều IB; không loại toàn bộ kết quả chỉ vì quantity mỗi dòng khác tổng yêu cầu.
- Chỉ trả “không tìm thấy trong phạm vi…” khi truy vấn đủ và đúng; thiếu quyền hoặc filter không áp dụng phải báo lỗi/partial.

### D. CHECK_PO — kiểm PO/SA trước nhận hàng

Ví dụ: `SAP kiểm tra PO <PO> item <item> mã <material> số lượng <qty> <unit>`.

- Lọc đúng PO/SA, item và material. Tách chế độ truy vấn khỏi selection parameter/layout đang dùng cho STO.
- Trả các quantity và trạng thái nguồn, chỉ suy số còn nhận khi đã xác minh đúng cột/đơn vị/kiểu chứng từ. Không gọi MGLIEF là “đã xuất” hoặc “còn GR” một cách cố định cho mọi layout.
- Hiển thị release, deletion/block và delivery complete nếu đọc được; giải thích theo cấu hình hệ thống đích.
- Một material nhiều PO item phải tách dòng; gợi ý phân bổ quantity có thể là kết quả tính toán riêng, không tự sửa PO hay tự tạo IB.
- Xác minh SA có đọc được qua ME2N/layout hiện tại; nếu chưa hỗ trợ, báo rõ thay vì tuyên bố kiểm PO/SA hoàn chỉnh.

### E. QUEUE_IB — tra queue lỗi liên quan inbound

Ví dụ: `SAP kiểm tra queue inbound <IB> kho <warehouse>`.

- Tách adapter queue monitor khỏi adapter warehouse-task/bin history dù cùng /SCWM/MON.
- Trả business object, error message, trạng thái queue nếu có và thời điểm/phạm vi kiểm tra.
- Warehouse/monitor/node là cấu hình đã xác minh, không hardcode warehouse của file mẫu cho mọi máy.
- Đọc đủ tất cả dòng, kể cả bảng chỉ có một dòng. Không chỉ nhận IB bắt đầu bằng 180 nếu chưa có quy tắc được xác nhận.
- Chỉ quan sát lỗi; không clear/restart queue và không thay queue của warehouse order.

### F. SHIPMENT_INFO — thông tin shipment/bill từ ZMME240

Ví dụ: `SAP tra shipment bill <reference> mã <material>`.

- Bật theo cấu hình khi SAP đích có ZMME240 và quyền báo cáo phù hợp.
- Trả nguyên mã nguồn cùng các cột đã xác minh: bill, PO/SA, material, quantity/unit, shipment status, mã liên quan inbound.
- Giữ `TRXID1` là mã nguồn cho đến khi xác minh nó là IB/reference nào. Không tự ánh xạ mọi shipment status thành trạng thái GR.
- Không tự fallback khi thiếu quyền để rồi kết luận không có shipment.

### G. PRECHECK_GR — báo cáo tổng hợp, triển khai sau các lệnh độc lập

Ví dụ: `SAP kiểm tra trước GR inbound <IB>`.

Ghép CHECK_IB, CHECK_PO, QUEUE_IB và TIM_GR_DOC theo cùng snapshot/phạm vi có timestamp từng nguồn. Báo các điều kiện đã quan sát, điểm thiếu và mâu thuẫn. Nếu một nguồn lỗi, giữ kết quả các nguồn còn lại và ghi chưa đầy đủ; không trả “sẵn sàng GR” chỉ vì quantity khớp hoặc không tìm thấy queue.

Đây là báo cáo hỗ trợ kiểm tra, không gọi Post_GR_EWM, MIGO post, đổi batch/SLoc hoặc sửa inbound phía sau.

## 4. Các thao tác ghi cần phạm vi riêng

| Nhóm | Macro gốc | Tác động quan sát được |
|---|---|---|
| GR thật | `Post_GR_EWM`, `GR_EWM_Combined`, `GR_MIGO_PO` | Ghi GR qua /SCWM/GR hoặc MIGO; có nút post/save thật |
| Header chứng từ | `Change_doc` | Sửa header text material document và lưu |
| Batch | `Create_Batch`, `Change_Batch`, `Delete_Batch` | Tạo/sửa HSD/Last GR Date/classification hoặc đặt deletion flags |
| Bin | `Create_BIN` | Tạo storage bin qua /SCWM/LS01 |
| Inbound | `Create_Inb_VL31n`, `Change_Inb`, `Change_Inb_Multi`, `Delete_Inb` | Tạo/sửa/xóa; có thay đổi SLoc/batch/date/quantity/item |
| WO queue | `Change_Queue_EWM` | Đổi queue warehouse order, không phải đọc queue lỗi |
| Outbound | `Complete_Outbound` | Sửa quantity, tạo FD và lưu trong /SCWM/PRDO |
| Reservation | `Delete_Res`, `Complete_Res` | Xóa hoặc đặt final issue qua MB22 |

Nếu triển khai các tác vụ này về sau, cần thiết kế riêng về quyền người gửi, preview từng dòng trước/sau, xác nhận thao tác cụ thể, lưu bằng chứng và phục hồi sau timeout. Retry webhook chỉ gửi lại kết quả, tuyệt đối không được tự gọi lại tác vụ ghi SAP. Chưa đưa các tác vụ này vào bot tra cứu trong đợt mở rộng đề xuất.

Các thủ tục menu `Batch_Process`, `GR_Doc_Process`, `Inbound_Process`, `GR_Process`, `Inb_Process` trộn đọc/ghi qua hộp thoại Yes/No; không chuyển nguyên wrapper thành một intent mơ hồ.

`Search_and_MIGO` có điền dữ liệu vào màn hình MIGO rồi bấm check; lệnh post tại dòng 2642 đang comment. Không mô tả nó là đã post chỉ vì thông báo cuối thủ tục ghi hoàn tất, cũng không coi đó là truy vấn chỉ đọc. `Post_GR_EWM1` là đoạn thao tác checkbox thiếu ngữ cảnh, không đưa vào roadmap.

## 5. Lỗi và giả định không được kế thừa

1. `Get_Mat_Doc` dùng `Right(XBLNR, 9)`, thêm 0 khi tìm reference và đọc dòng k-1 khi k=0. Cần mapping định danh rõ, không cắt chuỗi và không bỏ chứng từ bổ sung/reversal.
2. `Check_Inbound_Qty` ghép quantity vào khóa tìm kiếm, khiến chênh SL có thể bị báo thành sai material/IB. Thiết kế mới phải trả expected/actual/delta.
3. `Check_Queued` đặt rows=RowCount-1 rồi lặp đến rows-1: bỏ dòng cuối và không xử lý bảng một dòng. Không lấy “No queue” làm mặc định nếu đọc lỗi/thiếu quyền.
4. Các vòng FindNext thiếu kiểm tra quay lại firstAddress, có thể lặp vô hạn. Ghép dữ liệu bằng index/key có giới hạn và trả rõ trùng khóa.
5. Chuyển số theo dấu chấm/phẩy và CInt có thể mất phần lẻ/overflow. Dùng số numeric/Decimal và quy tắc locale của đợt sửa tính đúng.
6. Release code, layout tên cá nhân, warehouse/node monitor và mapping process type là giả định môi trường cũ. Không áp sang hệ thống khác khi chưa xác minh.
7. Module hỗ trợ có thông tin đăng nhập ghi trực tiếp trong source; không sao chép giá trị sang bridge, test, log hay gói phát hành. Chưa kiểm tra hiệu lực thông tin này. Bridge nên dùng phiên đã đăng nhập riêng, không dùng auto-login từ macro.
8. Get_SAPSession lấy connection/session đầu tiên; không đáp ứng yêu cầu pin phiên riêng. Workbook còn tự Save khi đóng, nên đọc bằng parser tĩnh như lần phân tích này để tránh làm thay đổi file.
9. Cột Excel thay nghĩa theo thủ tục: B có thể là inbound hoặc plant, F có thể là material document hoặc warehouse, E có thể là quantity hoặc queue. Không import toàn sheet bằng một schema cố định suy từ vị trí cột.

## 6. Thay đổi kiến trúc tối thiểu trong bridge

- Tạo bảng command registry chung cho canonical command, alias, validation, handler và output. Tránh allowlist JSON nhận lệnh mà main không có nhánh.
- Entity có kiểu: `materials`, `inboundNos`, `poNo`, `poItem`, `materialDocument`, `fiscalYear`, `reference`, `billNo`, `warehouse`, `expectedQty`, `unit`, `dateFrom/dateTo`. PO/IB/chứng từ không được đưa vào resolver vật tư.
- Giữ định danh dạng chuỗi và số 0 đầu. Khi số không rõ là vật tư/PO/IB, dùng nhãn trong câu hoặc hỏi rõ, không tự đoán và thao tác.
- Cùng yêu cầu JSON và tiếng Việt phải tạo cùng canonical metadata. Sửa parity hiện tại: JSON CHECK_OB/LICHSU_BIN chưa giữ đủ obNo/bin trong khi nhánh natural có.
- Câu “ngày/giờ GR của stock hiện tại” tiếp tục thuộc VITRI; “chứng từ/lịch sử GR của inbound/PO” đi TIM_GR_DOC. Test thứ tự intent để không phá câu đang dùng.
- Mở rộng có giới hạn adapter MB51/VL06I/ME2N; thêm adapter queue và Z-report riêng khi cần. Không nối toàn bộ VBA 3.509 dòng vào sap_engine.py.
- Kết quả dùng model numeric, source fields, scope và completeness của kế hoạch nâng cấp chính. Trường thiếu phải là unavailable, không phải 0 hoặc “chưa GR”.
- Giữ payload tương thích, Outbox và cách xưng hô hiện hành. Giới hạn byte phản hồi, cho phép nhiều kết quả/trích bảng có thông báo; không in dữ liệu nhạy cảm không cần vào nhóm.

## 7. Lộ trình bổ sung vào kế hoạch nâng cấp

| Giai đoạn | Phạm vi | Điều kiện trước khi bắt đầu |
|---|---|---|
| Nền tảng | Đợt 1 và 2 trong kế hoạch chính: số liệu, completeness, điều phối, Outbox và session | Có test hồi quy và môi trường thử riêng |
| Mở rộng A | TIM_GR_DOC, CHECK_IB, TIM_IB, CHECK_PO | Phạm vi query và cột SAP được xác minh; không còn fake-zero/numeric bug |
| Mở rộng B | QUEUE_IB, SHIPMENT_INFO và PRECHECK_GR | A ổn định; monitor/Z-report có quyền và dữ liệu mẫu biết trước |
| Khảo sát thêm | ZEWMR137 | Có export mẫu và định nghĩa cột/filter/KPI; chưa triển khai từ tên báo cáo |
| Tự động ghi | GR/create/change/delete và complete | Đề xuất riêng, không phụ thuộc việc bot tra cứu đã hoạt động |

Đợt tối ưu runtime/đóng gói có thể thực hiện riêng sau nền tảng, không trộn với thay đổi quy tắc GR. Chưa ấn định phiên bản phát hành cho các lệnh mới trước khi hoàn tất kiểm tra dữ liệu thật.

## 8. Tiêu chí kiểm thử trước phát hành lệnh mới

- Mỗi lệnh có fixture và test parse → validate → handler → numeric/completeness → HTML/Outbox; test JSON/natural cùng entity và lệnh thiếu tham số.
- IB nhiều item cùng material, một PO giao nhiều IB, nhiều PO item, quantity lệch, batch khác, UoM khác, mã có zero đầu.
- Một reference có nhiều GR, chứng từ khác năm, reversal/return và posting date khác document date; không lấy chứng từ đầu tiên làm toàn bộ lịch sử.
- Queue 0/1/n dòng, dòng cuối có lỗi, IB không bắt đầu 180, thiếu quyền/node và query timeout; báo phạm vi đầy đủ hoặc chưa đủ.
- ZMME240 không tồn tại/không có quyền, shipment có status lạ hoặc TRXID1 chưa xác định; giữ nguyên dữ liệu nguồn và giới hạn giải thích.
- Query bị cắt dòng, thiếu cột, không nhập đủ multi-select, ngày quá rộng và bộ lọc từ lần trước; không đưa kết luận nghiệp vụ khi prerequisites lỗi.
- Mock phải xác nhận các handler tra cứu không gọi post/save/change/delete của SAP. Kiểm thử thực tế dùng session riêng và account phù hợp, đối chiếu từng dòng với nguồn biết trước.
- Regression các lệnh cũ, giúp đỡ/PING, flow trả lời “Sếp”, payload budget, mất mạng/restart và gửi lại kết quả; không thay trigger @mention/SAP trong cùng đợt nếu chưa kiểm tra flow nhận.

## 9. Tác động đến dung lượng và Excel/CSV

Workbook gốc có vùng khai báo sheet GR_EWM_V1 đến hàng 1.032.367; đây không phải bằng chứng có hơn một triệu giao dịch. XML sheet1 khoảng 111,3 MB trước nén và 9,30 MB sau nén, chiếm phần lớn file 9,46 MB.

Nếu tiếp tục duy trì workbook, có thể khảo sát bản sao để xác định dữ liệu/style/row records thừa trước khi thu nhỏ; phải bảo toàn macro, drawing, nút bấm, công thức và named ranges. Không xóa hàng hoặc lưu lại workbook gốc chỉ dựa vào dimension. Lần phân tích này không sửa workbook.

Với bridge, không cần phân phối file này để chạy lệnh mới. Chỉ chuyển quy tắc đã xác minh và adapter cần thiết sang Python; giữ master vật tư hiện hành. Chưa có lý do đổi master sang CSV chỉ vì thêm chức năng GR; openpyxl vẫn phục vụ đọc SAP export.

Đầu ra lần này là tài liệu phân tích và phần bổ sung kế hoạch. Không có macro hoặc lệnh SAP mới nào đã được kích hoạt.
