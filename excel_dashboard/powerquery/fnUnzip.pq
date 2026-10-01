// fnUnzip : ZIP 파일(binary)의 첫 번째 항목을 압축 해제하여 binary로 반환
//  - KIS 종목마스터(kospi_code.mst.zip 등)는 파일 1개짜리 Deflate ZIP입니다.
//  - 중앙 디렉터리(EOCD)를 우선 읽고, 없으면 로컬 헤더 정보를 사용합니다.
(zip as binary) as binary =>
let
    B = Binary.Buffer(zip),
    Len = Binary.Length(B),
    U16 = (off as number) as number =>
        BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian)(Binary.Range(B, off, 2)),
    U32 = (off as number) as number =>
        BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian)(Binary.Range(B, off, 4)),
    LocalSig = U32(0),
    Eocd = Len - 22,
    HasEocd = Eocd > 0 and U32(Eocd) = 101010256,            // 0x06054b50
    CdOff = if HasEocd then U32(Eocd + 16) else null,
    FromCd = CdOff <> null and U32(CdOff) = 33639248,         // 0x02014b50
    CSize = if FromCd then U32(CdOff + 20) else U32(18),
    Method = if FromCd then U16(CdOff + 10) else U16(8),
    LhOff = if FromCd then U32(CdOff + 42) else 0,
    DataOff = LhOff + 30 + U16(LhOff + 26) + U16(LhOff + 28),
    Data = Binary.Range(B, DataOff, CSize),
    Result =
        if LocalSig <> 67324752 then error "ZIP 형식이 아닙니다."          // 0x04034b50
        else if Method = 0 then Data
        else if Method = 8 then Binary.Decompress(Data, Compression.Deflate)
        else error "지원하지 않는 ZIP 압축 방식입니다: " & Text.From(Method)
in
    Result
