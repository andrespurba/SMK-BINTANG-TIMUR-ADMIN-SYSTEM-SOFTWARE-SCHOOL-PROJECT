VERSION 5.00
Begin VB.UserControl jcbutton
  ScaleMode = 3
  AutoRedraw = True
  FontTransparent = True
  ClientLeft = 0
  ClientTop = 0
  ClientWidth = 14235
  ClientHeight = 8835
  BeginProperty Font
    Name = "Tahoma"
    Size = 8.25
    Charset = 0
    Weight = 400
    Underline = 0 'False
    Italic = 0 'False
    Strikethrough = 0 'False
  EndProperty
  Begin VB.Menu vsvs
    Caption = "vsd"
  End
End

Attribute VB_Name = "jcbutton"

Private Type UDT_1_006DD15C
  bStruc(16) As Byte ' String fields: 0
End Type

Private Type UDT_2_006DD168
  bStruc(28) As Byte ' String fields: 2
End Type

Private Type UDT_3_006DD188
  bStruc(20) As Byte ' String fields: 0
End Type

Private Type UDT_4_006DD15C
  bStruc(16) As Byte ' String fields: 0
End Type

Private Type UDT_5_006DD194
  bStruc(44) As Byte ' String fields: 1
End Type

Private Type UDT_6_006DD1A4
  bStruc(44) As Byte ' String fields: 0
End Type

Private Type UDT_7_006DD1B0
  bStruc(8) As Byte ' String fields: 0
End Type

Private Type UDT_8_006DD1BC
  bStruc(92) As Byte ' String fields: 1
End Type

Private Type UDT_9_006DD1D4
  bStruc(12) As Byte ' String fields: 0
End Type

Private Type UDT_10_006DD1E0
  bStruc(24) As Byte ' String fields: 0
End Type

Private Type UDT_11_006DD1EC
  bStruc(40) As Byte ' String fields: 0
End Type

Private Type UDT_12_006DD188
  bStruc(20) As Byte ' String fields: 0
End Type

Private Type UDT_13_006DD1F8
  bStruc(3) As Byte ' String fields: 0
End Type

Private Type UDT_14_006DD204
  bStruc(4) As Byte ' String fields: 0
End Type

Private Type UDT_15_006DD1A4
  bStruc(44) As Byte ' String fields: 0
End Type

Private Type UDT_16_006DD218
  bStruc(276) As Byte ' String fields: 1
End Type


Private Sub UserControl_UnknownEvent_7 'DFC420
  loc_00DFC420: push ebp
  loc_00DFC421: mov ebp, esp
  loc_00DFC423: sub esp, 0000000Ch
  loc_00DFC426: push 00402836h ; __vbaExceptHandler
  loc_00DFC42B: mov eax, fs:[00000000h]
  loc_00DFC431: push eax
  loc_00DFC432: mov fs:[00000000h], esp
  loc_00DFC439: sub esp, 00000010h
  loc_00DFC43C: push ebx
  loc_00DFC43D: push esi
  loc_00DFC43E: push edi
  loc_00DFC43F: mov var_C, esp
  loc_00DFC442: mov var_8, 00401940h
  loc_00DFC449: mov esi, Me
  loc_00DFC44C: mov eax, esi
  loc_00DFC44E: and eax, 00000001h
  loc_00DFC451: mov var_4, eax
  loc_00DFC454: and esi, FFFFFFFEh
  loc_00DFC457: push esi
  loc_00DFC458: mov Me, esi
  loc_00DFC45B: mov ecx, [esi]
  loc_00DFC45D: call [ecx+00000004h]
  loc_00DFC460: mov edx, [esi]
  loc_00DFC462: lea eax, var_18
  loc_00DFC465: xor ebx, ebx
  loc_00DFC467: push eax
  loc_00DFC468: push esi
  loc_00DFC469: mov var_18, ebx
  loc_00DFC46C: call [edx+00000088h]
  loc_00DFC472: cmp eax, ebx
  loc_00DFC474: fnclex
  loc_00DFC476: jge 00DFC48Eh
  loc_00DFC478: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC47E: push 00000088h
  loc_00DFC483: push 006DCB00h
  loc_00DFC488: push esi
  loc_00DFC489: push eax
  loc_00DFC48A: call edi
  loc_00DFC48C: jmp 00DFC494h
  loc_00DFC48E: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC494: fld real4 ptr var_18
  loc_00DFC497: fcomp real4 ptr [00401938h]
  loc_00DFC49D: fnstsw ax
  loc_00DFC49F: test ah, 01h
  loc_00DFC4A2: jz 00DFC4C6h
  loc_00DFC4A4: mov ecx, [esi]
  loc_00DFC4A6: push 435C0000h
  loc_00DFC4AB: push esi
  loc_00DFC4AC: call [ecx+0000008Ch]
  loc_00DFC4B2: cmp eax, ebx
  loc_00DFC4B4: fnclex
  loc_00DFC4B6: jge 00DFC4C6h
  loc_00DFC4B8: push 0000008Ch
  loc_00DFC4BD: push 006DCB00h
  loc_00DFC4C2: push esi
  loc_00DFC4C3: push eax
  loc_00DFC4C4: call edi
  loc_00DFC4C6: mov edx, [esi]
  loc_00DFC4C8: lea eax, var_18
  loc_00DFC4CB: push eax
  loc_00DFC4CC: push esi
  loc_00DFC4CD: call [edx+00000080h]
  loc_00DFC4D3: cmp eax, ebx
  loc_00DFC4D5: fnclex
  loc_00DFC4D7: jge 00DFC4E7h
  loc_00DFC4D9: push 00000080h
  loc_00DFC4DE: push 006DCB00h
  loc_00DFC4E3: push esi
  loc_00DFC4E4: push eax
  loc_00DFC4E5: call edi
  loc_00DFC4E7: fld real4 ptr var_18
  loc_00DFC4EA: fcomp real4 ptr [00401938h]
  loc_00DFC4F0: fnstsw ax
  loc_00DFC4F2: test ah, 01h
  loc_00DFC4F5: jz 00DFC519h
  loc_00DFC4F7: mov ecx, [esi]
  loc_00DFC4F9: push 435C0000h
  loc_00DFC4FE: push esi
  loc_00DFC4FF: call [ecx+00000084h]
  loc_00DFC505: cmp eax, ebx
  loc_00DFC507: fnclex
  loc_00DFC509: jge 00DFC519h
  loc_00DFC50B: push 00000084h
  loc_00DFC510: push 006DCB00h
  loc_00DFC515: push esi
  loc_00DFC516: push eax
  loc_00DFC517: call edi
  loc_00DFC519: mov edx, [esi]
  loc_00DFC51B: push esi
  loc_00DFC51C: call [edx+00000914h]
  loc_00DFC522: mov eax, [esi]
  loc_00DFC524: push esi
  loc_00DFC525: call [eax+00000910h]
  loc_00DFC52B: mov var_4, ebx
  loc_00DFC52E: mov eax, Me
  loc_00DFC531: push eax
  loc_00DFC532: mov ecx, [eax]
  loc_00DFC534: call [ecx+00000008h]
  loc_00DFC537: mov eax, var_4
  loc_00DFC53A: mov ecx, var_14
  loc_00DFC53D: pop edi
  loc_00DFC53E: pop esi
  loc_00DFC53F: mov fs:[00000000h], ecx
  loc_00DFC546: pop ebx
  loc_00DFC547: mov esp, ebp
  loc_00DFC549: pop ebp
  loc_00DFC54A: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_D 'DFA640
  loc_00DFA640: push ebp
  loc_00DFA641: mov ebp, esp
  loc_00DFA643: sub esp, 0000000Ch
  loc_00DFA646: push 00402836h ; __vbaExceptHandler
  loc_00DFA64B: mov eax, fs:[00000000h]
  loc_00DFA651: push eax
  loc_00DFA652: mov fs:[00000000h], esp
  loc_00DFA659: sub esp, 00000010h
  loc_00DFA65C: push ebx
  loc_00DFA65D: push esi
  loc_00DFA65E: push edi
  loc_00DFA65F: mov var_C, esp
  loc_00DFA662: mov var_8, 004018A8h
  loc_00DFA669: mov esi, Me
  loc_00DFA66C: mov eax, esi
  loc_00DFA66E: and eax, 00000001h
  loc_00DFA671: mov var_4, eax
  loc_00DFA674: and esi, FFFFFFFEh
  loc_00DFA677: push esi
  loc_00DFA678: mov Me, esi
  loc_00DFA67B: mov ecx, [esi]
  loc_00DFA67D: call [ecx+00000004h]
  loc_00DFA680: xor ebx, ebx
  loc_00DFA682: cmp [esi+00000052h], bx
  loc_00DFA686: mov var_18, ebx
  loc_00DFA689: jz 00DFA69Ah
  loc_00DFA68B: mov edx, [esi+00000054h]
  loc_00DFA68E: push edx
  loc_00DFA68F: call 006DDD70h ; SetCursor(%x1v)
  loc_00DFA694: call [00401074h] ; __vbaSetSystemError
  loc_00DFA69A: cmp [esi+000000D4h], 0001h
  loc_00DFA6A2: lea edi, [esi+000000D4h]
  loc_00DFA6A8: jnz 00DFA75Ah
  loc_00DFA6AE: mov eax, [esi]
  loc_00DFA6B0: lea ecx, var_18
  loc_00DFA6B3: push ecx
  loc_00DFA6B4: push esi
  loc_00DFA6B5: call [eax+00000818h]
  loc_00DFA6BB: cmp eax, ebx
  loc_00DFA6BD: jge 00DFA6D1h
  loc_00DFA6BF: push 00000818h
  loc_00DFA6C4: push 006DD090h
  loc_00DFA6C9: push esi
  loc_00DFA6CA: push eax
  loc_00DFA6CB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFA6D1: mov edx, var_18
  loc_00DFA6D4: push edx
  loc_00DFA6D5: call 006DE1D8h ; SetCapture(%x1v)
  loc_00DFA6DA: call [00401074h] ; __vbaSetSystemError
  loc_00DFA6E0: mov ecx, [esi+00000048h]
  loc_00DFA6E3: mov eax, 00000002h
  loc_00DFA6E8: cmp ecx, eax
  loc_00DFA6EA: jz 00DFA6EFh
  loc_00DFA6EC: mov [esi+00000048h], eax
  loc_00DFA6EF: mov eax, [esi]
  loc_00DFA6F1: push esi
  loc_00DFA6F2: call [eax+00000910h]
  loc_00DFA6F8: mov ecx, [esi]
  loc_00DFA6FA: lea edx, [esi+000000DCh]
  loc_00DFA700: push edx
  loc_00DFA701: lea eax, [esi+000000D8h]
  loc_00DFA707: lea edx, [esi+000000D6h]
  loc_00DFA70D: push eax
  loc_00DFA70E: push edx
  loc_00DFA70F: push edi
  loc_00DFA710: push esi
  loc_00DFA711: call [ecx+000009ACh]
  loc_00DFA717: cmp eax, ebx
  loc_00DFA719: jge 00DFA72Dh
  loc_00DFA71B: push 000009ACh
  loc_00DFA720: push 006DD090h
  loc_00DFA725: push esi
  loc_00DFA726: push eax
  loc_00DFA727: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFA72D: cmp [esi+000000E0h], bx
  loc_00DFA734: jnz 00DFA748h
  loc_00DFA736: push ebx
  loc_00DFA737: push FFFFFDA7h
  loc_00DFA73C: push esi
  loc_00DFA73D: call [00401044h] ; __vbaRaiseEvent
  loc_00DFA743: add esp, 0000000Ch
  loc_00DFA746: jmp 00DFA75Ah
  loc_00DFA748: cmp [esi+000000E2h], bx
  loc_00DFA74F: jnz 00DFA75Ah
  loc_00DFA751: mov eax, [esi]
  loc_00DFA753: push esi
  loc_00DFA754: call [eax+00000978h]
  loc_00DFA75A: mov var_4, ebx
  loc_00DFA75D: mov eax, Me
  loc_00DFA760: push eax
  loc_00DFA761: mov ecx, [eax]
  loc_00DFA763: call [ecx+00000008h]
  loc_00DFA766: mov eax, var_4
  loc_00DFA769: mov ecx, var_14
  loc_00DFA76C: pop edi
  loc_00DFA76D: pop esi
  loc_00DFA76E: mov fs:[00000000h], ecx
  loc_00DFA775: pop ebx
  loc_00DFA776: mov esp, ebp
  loc_00DFA778: pop ebp
  loc_00DFA779: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_E 'DFA780
  loc_00DFA780: push ebp
  loc_00DFA781: mov ebp, esp
  loc_00DFA783: sub esp, 0000000Ch
  loc_00DFA786: push 00402836h ; __vbaExceptHandler
  loc_00DFA78B: mov eax, fs:[00000000h]
  loc_00DFA791: push eax
  loc_00DFA792: mov fs:[00000000h], esp
  loc_00DFA799: sub esp, 00000008h
  loc_00DFA79C: push ebx
  loc_00DFA79D: push esi
  loc_00DFA79E: push edi
  loc_00DFA79F: mov var_C, esp
  loc_00DFA7A2: mov var_8, 004018B0h
  loc_00DFA7A9: mov esi, Me
  loc_00DFA7AC: mov eax, esi
  loc_00DFA7AE: and eax, 00000001h
  loc_00DFA7B1: mov var_4, eax
  loc_00DFA7B4: and esi, FFFFFFFEh
  loc_00DFA7B7: push esi
  loc_00DFA7B8: mov Me, esi
  loc_00DFA7BB: mov ecx, [esi]
  loc_00DFA7BD: call [ecx+00000004h]
  loc_00DFA7C0: mov [esi+00000050h], FFFFFFh
  loc_00DFA7C6: mov var_4, 00000000h
  loc_00DFA7CD: mov eax, Me
  loc_00DFA7D0: push eax
  loc_00DFA7D1: mov edx, [eax]
  loc_00DFA7D3: call [edx+00000008h]
  loc_00DFA7D6: mov eax, var_4
  loc_00DFA7D9: mov ecx, var_14
  loc_00DFA7DC: pop edi
  loc_00DFA7DD: pop esi
  loc_00DFA7DE: mov fs:[00000000h], ecx
  loc_00DFA7E5: pop ebx
  loc_00DFA7E6: mov esp, ebp
  loc_00DFA7E8: pop ebp
  loc_00DFA7E9: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_F 'DFAD70
  loc_00DFAD70: push ebp
  loc_00DFAD71: mov ebp, esp
  loc_00DFAD73: sub esp, 0000000Ch
  loc_00DFAD76: push 00402836h ; __vbaExceptHandler
  loc_00DFAD7B: mov eax, fs:[00000000h]
  loc_00DFAD81: push eax
  loc_00DFAD82: mov fs:[00000000h], esp
  loc_00DFAD89: sub esp, 00000048h
  loc_00DFAD8C: push ebx
  loc_00DFAD8D: push esi
  loc_00DFAD8E: push edi
  loc_00DFAD8F: mov var_C, esp
  loc_00DFAD92: mov var_8, 004018E8h
  loc_00DFAD99: mov esi, Me
  loc_00DFAD9C: mov eax, esi
  loc_00DFAD9E: and eax, 00000001h
  loc_00DFADA1: mov var_4, eax
  loc_00DFADA4: and esi, FFFFFFFEh
  loc_00DFADA7: push esi
  loc_00DFADA8: mov Me, esi
  loc_00DFADAB: mov ecx, [esi]
  loc_00DFADAD: call [ecx+00000004h]
  loc_00DFADB0: mov edx, Cancel
  loc_00DFADB3: xor ebx, ebx
  loc_00DFADB5: mov var_24, ebx
  loc_00DFADB8: mov var_34, ebx
  loc_00DFADBB: movsx eax, [edx]
  loc_00DFADBE: add eax, FFFFFFF3h
  loc_00DFADC1: mov var_44, ebx
  loc_00DFADC4: cmp eax, 0000001Bh
  loc_00DFADC7: mov var_48, ebx
  loc_00DFADCA: mov var_4C, ebx
  loc_00DFADCD: ja 00DFAF53h
  loc_00DFADD3: xor ecx, ecx
  loc_00DFADD5: mov cl, [eax+00DFB004h]
  loc_00DFADDB: jmp [ecx*4+00DFAFF0h]
  loc_00DFADE2: push ebx
  loc_00DFADE3: push FFFFFDA8h
  loc_00DFADE8: push esi
  loc_00DFADE9: call [00401044h] ; __vbaRaiseEvent
  loc_00DFADEF: add esp, 0000000Ch
  loc_00DFADF2: jmp 00DFAF6Fh
  loc_00DFADF7: lea edx, var_24
  loc_00DFADFA: mov var_1C, 80020004h
  loc_00DFAE01: push edx
  loc_00DFAE02: push 006DEB2Ch ; "+{TAB}"
  loc_00DFAE07: mov var_24, 0000000Ah
  loc_00DFAE0E: call [004010C4h] ; rtcSendKeys
  loc_00DFAE14: lea ecx, var_24
  loc_00DFAE17: call [00401024h] ; __vbaFreeVar
  loc_00DFAE1D: jmp 00DFAF6Fh
  loc_00DFAE22: lea eax, var_24
  loc_00DFAE25: mov var_1C, 80020004h
  loc_00DFAE2C: push eax
  loc_00DFAE2D: push 006DEB40h ; "{TAB}"
  loc_00DFAE32: mov var_24, 0000000Ah
  loc_00DFAE39: call [004010C4h] ; rtcSendKeys
  loc_00DFAE3F: lea ecx, var_24
  loc_00DFAE42: call [00401024h] ; __vbaFreeVar
  loc_00DFAE48: jmp 00DFAF6Fh
  loc_00DFAE4D: mov ecx, arg_10
  loc_00DFAE50: cmp [ecx], 0004h
  loc_00DFAE54: jz 00DFAFB9h
  loc_00DFAE5A: cmp [esi+0000004Ch], bx
  loc_00DFAE5E: jnz 00DFAEABh
  loc_00DFAE60: mov eax, [esi+00000064h]
  loc_00DFAE63: or edi, FFFFFFFFh
  loc_00DFAE66: cmp eax, 00000001h
  loc_00DFAE69: mov [esi+000000A8h], di
  loc_00DFAE70: jnz 00DFAE7Fh
  loc_00DFAE72: mov dx, [esi+0000006Ch]
  loc_00DFAE76: not dx
  loc_00DFAE79: mov [esi+0000006Ch], dx
  loc_00DFAE7D: jmp 00DFAE91h
  loc_00DFAE7F: cmp eax, 00000002h
  loc_00DFAE82: jnz 00DFAE91h
  loc_00DFAE84: mov eax, [esi]
  loc_00DFAE86: push esi
  loc_00DFAE87: call [eax+00000938h]
  loc_00DFAE8D: mov [esi+0000006Ch], di
  loc_00DFAE91: mov ecx, [esi+00000048h]
  loc_00DFAE94: mov eax, 00000002h
  loc_00DFAE99: cmp ecx, eax
  loc_00DFAE9B: jz 00DFAEDBh
  loc_00DFAE9D: mov ecx, [esi]
  loc_00DFAE9F: push esi
  loc_00DFAEA0: mov [esi+00000048h], eax
  loc_00DFAEA3: call [ecx+00000910h]
  loc_00DFAEA9: jmp 00DFAEDBh
  loc_00DFAEAB: cmp [esi+0000004Eh], bx
  loc_00DFAEAF: mov eax, [esi+00000048h]
  loc_00DFAEB2: jz 00DFAECBh
  loc_00DFAEB4: cmp eax, 00000002h
  loc_00DFAEB7: jz 00DFAEDBh
  loc_00DFAEB9: mov edx, [esi]
  loc_00DFAEBB: push esi
  loc_00DFAEBC: mov [esi+00000048h], 00000002h
  loc_00DFAEC3: call [edx+00000910h]
  loc_00DFAEC9: jmp 00DFAEDBh
  loc_00DFAECB: cmp eax, ebx
  loc_00DFAECD: jz 00DFAEDBh
  loc_00DFAECF: mov eax, [esi]
  loc_00DFAED1: push esi
  loc_00DFAED2: mov [esi+00000048h], ebx
  loc_00DFAED5: call [eax+00000910h]
  loc_00DFAEDB: call 006DE234h ; GetCapture()
  loc_00DFAEE0: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DFAEE6: mov var_48, eax
  loc_00DFAEE9: call edi
  loc_00DFAEEB: mov eax, [esi+00000010h]
  loc_00DFAEEE: lea edx, var_4C
  loc_00DFAEF1: push edx
  loc_00DFAEF2: push eax
  loc_00DFAEF3: mov ecx, [eax]
  loc_00DFAEF5: call [ecx+00000058h]
  loc_00DFAEF8: cmp eax, ebx
  loc_00DFAEFA: fnclex
  loc_00DFAEFC: jge 00DFAF10h
  loc_00DFAEFE: mov ecx, [esi+00000010h]
  loc_00DFAF01: push 00000058h
  loc_00DFAF03: push 006DCB00h
  loc_00DFAF08: push ecx
  loc_00DFAF09: push eax
  loc_00DFAF0A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFAF10: mov edx, var_48
  loc_00DFAF13: mov eax, var_4C
  loc_00DFAF16: cmp edx, eax
  loc_00DFAF18: jz 00DFAF6Fh
  loc_00DFAF1A: call 006DE194h ; ReleaseCapture()
  loc_00DFAF1F: call edi
  loc_00DFAF21: mov eax, [esi+00000010h]
  loc_00DFAF24: lea edx, var_48
  loc_00DFAF27: push edx
  loc_00DFAF28: push eax
  loc_00DFAF29: mov ecx, [eax]
  loc_00DFAF2B: call [ecx+00000058h]
  loc_00DFAF2E: cmp eax, ebx
  loc_00DFAF30: fnclex
  loc_00DFAF32: jge 00DFAF46h
  loc_00DFAF34: mov ecx, [esi+00000010h]
  loc_00DFAF37: push 00000058h
  loc_00DFAF39: push 006DCB00h
  loc_00DFAF3E: push ecx
  loc_00DFAF3F: push eax
  loc_00DFAF40: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFAF46: mov edx, var_48
  loc_00DFAF49: push edx
  loc_00DFAF4A: call 006DE1D8h ; SetCapture(%x1v)
  loc_00DFAF4F: call edi
  loc_00DFAF51: jmp 00DFAF6Fh
  loc_00DFAF53: cmp [esi+000000A8h], bx
  loc_00DFAF5A: jz 00DFAF6Fh
  loc_00DFAF5C: mov eax, [esi]
  loc_00DFAF5E: push esi
  loc_00DFAF5F: mov [esi+000000A8h], bx
  loc_00DFAF66: mov [esi+00000048h], ebx
  loc_00DFAF69: call [eax+00000910h]
  loc_00DFAF6F: sub esp, 00000010h
  loc_00DFAF72: mov eax, 00004002h
  loc_00DFAF77: mov edx, esp
  loc_00DFAF79: mov ecx, eax
  loc_00DFAF7B: sub esp, 00000010h
  loc_00DFAF7E: mov [edx], eax
  loc_00DFAF80: mov eax, var_30
  loc_00DFAF83: mov [edx+00000004h], eax
  loc_00DFAF86: mov eax, Cancel
  loc_00DFAF89: mov [edx+00000008h], eax
  loc_00DFAF8C: mov eax, var_28
  loc_00DFAF8F: mov [edx+0000000Ch], eax
  loc_00DFAF92: mov eax, var_40
  loc_00DFAF95: mov edx, esp
  loc_00DFAF97: push 00000002h
  loc_00DFAF99: push FFFFFDA6h
  loc_00DFAF9E: push esi
  loc_00DFAF9F: mov [edx], ecx
  loc_00DFAFA1: mov ecx, arg_10
  loc_00DFAFA4: mov [edx+00000004h], eax
  loc_00DFAFA7: mov eax, var_38
  loc_00DFAFAA: mov [edx+00000008h], ecx
  loc_00DFAFAD: mov [edx+0000000Ch], eax
  loc_00DFAFB0: call [00401044h] ; __vbaRaiseEvent
  loc_00DFAFB6: add esp, 0000002Ch
  loc_00DFAFB9: mov var_4, ebx
  loc_00DFAFBC: push 00DFAFCEh
  loc_00DFAFC1: jmp 00DFAFCDh
  loc_00DFAFC3: lea ecx, var_24
  loc_00DFAFC6: call [00401024h] ; __vbaFreeVar
  loc_00DFAFCC: ret
  loc_00DFAFCD: ret
  loc_00DFAFCE: mov eax, Me
  loc_00DFAFD1: push eax
  loc_00DFAFD2: mov ecx, [eax]
  loc_00DFAFD4: call [ecx+00000008h]
  loc_00DFAFD7: mov eax, var_4
  loc_00DFAFDA: mov ecx, var_14
  loc_00DFAFDD: pop edi
  loc_00DFAFDE: pop esi
  loc_00DFAFDF: mov fs:[00000000h], ecx
  loc_00DFAFE6: pop ebx
  loc_00DFAFE7: mov esp, ebp
  loc_00DFAFE9: pop ebp
  loc_00DFAFEA: retn 000Ch
End Sub

Private Sub UserControl_UnknownEvent_10 'DFB020
  loc_00DFB020: push ebp
  loc_00DFB021: mov ebp, esp
  loc_00DFB023: sub esp, 0000000Ch
  loc_00DFB026: push 00402836h ; __vbaExceptHandler
  loc_00DFB02B: mov eax, fs:[00000000h]
  loc_00DFB031: push eax
  loc_00DFB032: mov fs:[00000000h], esp
  loc_00DFB039: sub esp, 00000018h
  loc_00DFB03C: push ebx
  loc_00DFB03D: push esi
  loc_00DFB03E: push edi
  loc_00DFB03F: mov var_C, esp
  loc_00DFB042: mov var_8, 004018F8h
  loc_00DFB049: mov esi, Me
  loc_00DFB04C: mov eax, esi
  loc_00DFB04E: and eax, 00000001h
  loc_00DFB051: mov var_4, eax
  loc_00DFB054: and esi, FFFFFFFEh
  loc_00DFB057: push esi
  loc_00DFB058: mov Me, esi
  loc_00DFB05B: mov ecx, [esi]
  loc_00DFB05D: call [ecx+00000004h]
  loc_00DFB060: sub esp, 00000010h
  loc_00DFB063: mov eax, Cancel
  loc_00DFB066: mov edx, esp
  loc_00DFB068: mov ecx, 00004002h
  loc_00DFB06D: push 00000001h
  loc_00DFB06F: push FFFFFDA5h
  loc_00DFB074: mov [edx], ecx
  loc_00DFB076: mov ecx, var_20
  loc_00DFB079: push esi
  loc_00DFB07A: mov [edx+00000004h], ecx
  loc_00DFB07D: mov [edx+00000008h], eax
  loc_00DFB080: mov eax, var_18
  loc_00DFB083: mov [edx+0000000Ch], eax
  loc_00DFB086: call [00401044h] ; __vbaRaiseEvent
  loc_00DFB08C: add esp, 0000001Ch
  loc_00DFB08F: mov var_4, 00000000h
  loc_00DFB096: mov eax, Me
  loc_00DFB099: push eax
  loc_00DFB09A: mov ecx, [eax]
  loc_00DFB09C: call [ecx+00000008h]
  loc_00DFB09F: mov eax, var_4
  loc_00DFB0A2: mov ecx, var_14
  loc_00DFB0A5: pop edi
  loc_00DFB0A6: pop esi
  loc_00DFB0A7: mov fs:[00000000h], ecx
  loc_00DFB0AE: pop ebx
  loc_00DFB0AF: mov esp, ebp
  loc_00DFB0B1: pop ebp
  loc_00DFB0B2: retn 0008h
End Sub

Private Sub UserControl_UnknownEvent_11 'DFB0C0
  loc_00DFB0C0: push ebp
  loc_00DFB0C1: mov ebp, esp
  loc_00DFB0C3: sub esp, 0000000Ch
  loc_00DFB0C6: push 00402836h ; __vbaExceptHandler
  loc_00DFB0CB: mov eax, fs:[00000000h]
  loc_00DFB0D1: push eax
  loc_00DFB0D2: mov fs:[00000000h], esp
  loc_00DFB0D9: sub esp, 00000028h
  loc_00DFB0DC: push ebx
  loc_00DFB0DD: push esi
  loc_00DFB0DE: push edi
  loc_00DFB0DF: mov var_C, esp
  loc_00DFB0E2: mov var_8, 00401900h
  loc_00DFB0E9: mov esi, Me
  loc_00DFB0EC: mov ebx, 00000001h
  loc_00DFB0F1: mov eax, esi
  loc_00DFB0F3: and eax, ebx
  loc_00DFB0F5: mov var_4, eax
  loc_00DFB0F8: and esi, FFFFFFFEh
  loc_00DFB0FB: push esi
  loc_00DFB0FC: mov Me, esi
  loc_00DFB0FF: mov ecx, [esi]
  loc_00DFB101: call [ecx+00000004h]
  loc_00DFB104: mov edx, Cancel
  loc_00DFB107: xor edi, edi
  loc_00DFB109: mov var_24, edi
  loc_00DFB10C: mov var_34, edi
  loc_00DFB10F: cmp [edx], 0020h
  loc_00DFB113: jnz 00DFB1F9h
  loc_00DFB119: call 006DE194h ; ReleaseCapture()
  loc_00DFB11E: call [00401074h] ; __vbaSetSystemError
  loc_00DFB124: mov ax, [esi+0000004Eh]
  loc_00DFB128: cmp ax, di
  loc_00DFB12B: jz 00DFB174h
  loc_00DFB12D: cmp [esi+0000004Ch], di
  loc_00DFB131: jz 00DFB14Dh
  loc_00DFB133: mov ecx, [esi+00000048h]
  loc_00DFB136: mov eax, 00000002h
  loc_00DFB13B: cmp ecx, eax
  loc_00DFB13D: jz 00DFB1A4h
  loc_00DFB13F: mov [esi+00000048h], eax
  loc_00DFB142: mov eax, [esi]
  loc_00DFB144: push esi
  loc_00DFB145: call [eax+00000910h]
  loc_00DFB14B: jmp 00DFB1A4h
  loc_00DFB14D: cmp ax, di
  loc_00DFB150: jz 00DFB174h
  loc_00DFB152: cmp [esi+00000048h], ebx
  loc_00DFB155: jz 00DFB163h
  loc_00DFB157: mov ecx, [esi]
  loc_00DFB159: push esi
  loc_00DFB15A: mov [esi+00000048h], ebx
  loc_00DFB15D: call [ecx+00000910h]
  loc_00DFB163: cmp [esi+0000004Ch], di
  loc_00DFB167: jnz 00DFB1A4h
  loc_00DFB169: cmp [esi+000000A8h], di
  loc_00DFB170: jz 00DFB1A4h
  loc_00DFB172: jmp 00DFB194h
  loc_00DFB174: cmp [esi+00000048h], edi
  loc_00DFB177: jz 00DFB185h
  loc_00DFB179: mov edx, [esi]
  loc_00DFB17B: push esi
  loc_00DFB17C: mov [esi+00000048h], edi
  loc_00DFB17F: call [edx+00000910h]
  loc_00DFB185: cmp [esi+0000004Ch], di
  loc_00DFB189: jnz 00DFB1A4h
  loc_00DFB18B: cmp [esi+000000A8h], di
  loc_00DFB192: jz 00DFB1A4h
  loc_00DFB194: push edi
  loc_00DFB195: push FFFFFDA8h
  loc_00DFB19A: push esi
  loc_00DFB19B: call [00401044h] ; __vbaRaiseEvent
  loc_00DFB1A1: add esp, 0000000Ch
  loc_00DFB1A4: sub esp, 00000010h
  loc_00DFB1A7: mov eax, 00004002h
  loc_00DFB1AC: mov ebx, esp
  loc_00DFB1AE: mov edx, eax
  loc_00DFB1B0: sub esp, 00000010h
  loc_00DFB1B3: mov ecx, arg_10
  loc_00DFB1B6: mov [ebx], eax
  loc_00DFB1B8: mov eax, var_20
  loc_00DFB1BB: mov [ebx+00000004h], eax
  loc_00DFB1BE: mov eax, Cancel
  loc_00DFB1C1: mov [ebx+00000008h], eax
  loc_00DFB1C4: mov eax, var_18
  loc_00DFB1C7: mov [ebx+0000000Ch], eax
  loc_00DFB1CA: mov eax, esp
  loc_00DFB1CC: push 00000002h
  loc_00DFB1CE: push FFFFFDA4h
  loc_00DFB1D3: mov [eax], edx
  loc_00DFB1D5: mov edx, var_30
  loc_00DFB1D8: push esi
  loc_00DFB1D9: mov [eax+00000004h], edx
  loc_00DFB1DC: mov [eax+00000008h], ecx
  loc_00DFB1DF: mov ecx, var_28
  loc_00DFB1E2: mov [eax+0000000Ch], ecx
  loc_00DFB1E5: call [00401044h] ; __vbaRaiseEvent
  loc_00DFB1EB: add esp, 0000002Ch
  loc_00DFB1EE: mov [esi+000000A8h], di
  loc_00DFB1F5: mov [esi+0000004Ch], di
  loc_00DFB1F9: mov var_4, edi
  loc_00DFB1FC: mov eax, Me
  loc_00DFB1FF: push eax
  loc_00DFB200: mov edx, [eax]
  loc_00DFB202: call [edx+00000008h]
  loc_00DFB205: mov eax, var_4
  loc_00DFB208: mov ecx, var_14
  loc_00DFB20B: pop edi
  loc_00DFB20C: pop esi
  loc_00DFB20D: mov fs:[00000000h], ecx
  loc_00DFB214: pop ebx
  loc_00DFB215: mov esp, ebp
  loc_00DFB217: pop ebp
  loc_00DFB218: retn 000Ch
End Sub

Private Sub UserControl_UnknownEvent_12 'DFB220
  loc_00DFB220: push ebp
  loc_00DFB221: mov ebp, esp
  loc_00DFB223: sub esp, 0000000Ch
  loc_00DFB226: push 00402836h ; __vbaExceptHandler
  loc_00DFB22B: mov eax, fs:[00000000h]
  loc_00DFB231: push eax
  loc_00DFB232: mov fs:[00000000h], esp
  loc_00DFB239: sub esp, 00000008h
  loc_00DFB23C: push ebx
  loc_00DFB23D: push esi
  loc_00DFB23E: push edi
  loc_00DFB23F: mov var_C, esp
  loc_00DFB242: mov var_8, 00401908h
  loc_00DFB249: mov esi, Me
  loc_00DFB24C: mov ebx, 00000001h
  loc_00DFB251: mov eax, esi
  loc_00DFB253: and eax, ebx
  loc_00DFB255: mov var_4, eax
  loc_00DFB258: and esi, FFFFFFFEh
  loc_00DFB25B: push esi
  loc_00DFB25C: mov Me, esi
  loc_00DFB25F: mov ecx, [esi]
  loc_00DFB261: call [ecx+00000004h]
  loc_00DFB264: xor edi, edi
  loc_00DFB266: cmp [esi+00000070h], di
  loc_00DFB26A: mov [esi+00000050h], di
  loc_00DFB26E: mov [esi+0000004Ch], di
  loc_00DFB272: mov [esi+000000A8h], di
  loc_00DFB279: jnz 00DFB282h
  loc_00DFB27B: cmp [esi+00000048h], edi
  loc_00DFB27E: jz 00DFB29Bh
  loc_00DFB280: jmp 00DFB298h
  loc_00DFB282: cmp [esi+0000004Eh], di
  loc_00DFB286: mov eax, [esi+00000048h]
  loc_00DFB289: jz 00DFB294h
  loc_00DFB28B: cmp eax, ebx
  loc_00DFB28D: jz 00DFB29Bh
  loc_00DFB28F: mov [esi+00000048h], ebx
  loc_00DFB292: jmp 00DFB29Bh
  loc_00DFB294: cmp eax, edi
  loc_00DFB296: jz 00DFB29Bh
  loc_00DFB298: mov [esi+00000048h], edi
  loc_00DFB29B: mov edx, [esi]
  loc_00DFB29D: push esi
  loc_00DFB29E: call [edx+00000910h]
  loc_00DFB2A4: cmp [esi+00000058h], di
  loc_00DFB2A8: jz 00DFB2B3h
  loc_00DFB2AA: mov eax, [esi]
  loc_00DFB2AC: push esi
  loc_00DFB2AD: call [eax+00000910h]
  loc_00DFB2B3: mov var_4, edi
  loc_00DFB2B6: mov eax, Me
  loc_00DFB2B9: push eax
  loc_00DFB2BA: mov ecx, [eax]
  loc_00DFB2BC: call [ecx+00000008h]
  loc_00DFB2BF: mov eax, var_4
  loc_00DFB2C2: mov ecx, var_14
  loc_00DFB2C5: pop edi
  loc_00DFB2C6: pop esi
  loc_00DFB2C7: mov fs:[00000000h], ecx
  loc_00DFB2CE: pop ebx
  loc_00DFB2CF: mov esp, ebp
  loc_00DFB2D1: pop ebp
  loc_00DFB2D2: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_13 'DFB2E0
  loc_00DFB2E0: push ebp
  loc_00DFB2E1: mov ebp, esp
  loc_00DFB2E3: sub esp, 0000000Ch
  loc_00DFB2E6: push 00402836h ; __vbaExceptHandler
  loc_00DFB2EB: mov eax, fs:[00000000h]
  loc_00DFB2F1: push eax
  loc_00DFB2F2: mov fs:[00000000h], esp
  loc_00DFB2F9: sub esp, 00000048h
  loc_00DFB2FC: push ebx
  loc_00DFB2FD: push esi
  loc_00DFB2FE: push edi
  loc_00DFB2FF: mov var_C, esp
  loc_00DFB302: mov var_8, 00401910h
  loc_00DFB309: mov esi, Me
  loc_00DFB30C: mov eax, esi
  loc_00DFB30E: and eax, 00000001h
  loc_00DFB311: mov var_4, eax
  loc_00DFB314: and esi, FFFFFFFEh
  loc_00DFB317: push esi
  loc_00DFB318: mov Me, esi
  loc_00DFB31B: mov ecx, [esi]
  loc_00DFB31D: call [ecx+00000004h]
  loc_00DFB320: mov ebx, Cancel
  loc_00DFB323: mov eax, arg_14
  loc_00DFB326: xor edi, edi
  loc_00DFB328: mov dx, [ebx]
  loc_00DFB32B: cmp [esi+00000052h], di
  loc_00DFB32F: mov [esi+000000D4h], dx
  loc_00DFB336: mov ecx, [eax]
  loc_00DFB338: mov edx, arg_18
  loc_00DFB33B: mov [esi+000000D8h], ecx
  loc_00DFB341: mov ecx, arg_10
  loc_00DFB344: mov var_24, edi
  loc_00DFB347: mov eax, [edx]
  loc_00DFB349: mov var_34, edi
  loc_00DFB34C: mov [esi+000000DCh], eax
  loc_00DFB352: mov dx, [ecx]
  loc_00DFB355: mov var_44, edi
  loc_00DFB358: mov var_54, edi
  loc_00DFB35B: mov [esi+000000D6h], dx
  loc_00DFB362: jz 00DFB373h
  loc_00DFB364: mov eax, [esi+00000054h]
  loc_00DFB367: push eax
  loc_00DFB368: call 006DDD70h ; SetCursor(%x1v)
  loc_00DFB36D: call [00401074h] ; __vbaSetSystemError
  loc_00DFB373: cmp [ebx], 0001h
  loc_00DFB377: jz 00DFB386h
  loc_00DFB379: cmp [esi+000000E2h], di
  loc_00DFB380: jz 00DFB458h
  loc_00DFB386: or eax, FFFFFFFFh
  loc_00DFB389: cmp [esi+000000A8h], di
  loc_00DFB390: mov [esi+00000050h], ax
  loc_00DFB394: mov [esi+0000004Ch], ax
  loc_00DFB398: jnz 00DFB3B2h
  loc_00DFB39A: mov ecx, [esi+00000048h]
  loc_00DFB39D: mov eax, 00000002h
  loc_00DFB3A2: cmp ecx, eax
  loc_00DFB3A4: jz 00DFB3B2h
  loc_00DFB3A6: mov ecx, [esi]
  loc_00DFB3A8: push esi
  loc_00DFB3A9: mov [esi+00000048h], eax
  loc_00DFB3AC: call [ecx+00000910h]
  loc_00DFB3B2: cmp [esi+000000E0h], di
  loc_00DFB3B9: jnz 00DFB446h
  loc_00DFB3BF: sub esp, 00000010h
  loc_00DFB3C2: mov eax, 00004002h
  loc_00DFB3C7: mov ebx, esp
  loc_00DFB3C9: mov ecx, eax
  loc_00DFB3CB: sub esp, 00000010h
  loc_00DFB3CE: mov edx, 00004004h
  loc_00DFB3D3: mov [ebx], eax
  loc_00DFB3D5: mov eax, var_20
  loc_00DFB3D8: mov edi, edx
  loc_00DFB3DA: mov [ebx+00000004h], eax
  loc_00DFB3DD: mov eax, Cancel
  loc_00DFB3E0: mov [ebx+00000008h], eax
  loc_00DFB3E3: mov eax, var_18
  loc_00DFB3E6: mov [ebx+0000000Ch], eax
  loc_00DFB3E9: mov eax, esp
  loc_00DFB3EB: sub esp, 00000010h
  loc_00DFB3EE: mov [eax], ecx
  loc_00DFB3F0: mov ecx, var_30
  loc_00DFB3F3: mov [eax+00000004h], ecx
  loc_00DFB3F6: mov ecx, arg_10
  loc_00DFB3F9: mov [eax+00000008h], ecx
  loc_00DFB3FC: mov ecx, var_28
  loc_00DFB3FF: mov [eax+0000000Ch], ecx
  loc_00DFB402: mov ecx, var_40
  loc_00DFB405: mov eax, esp
  loc_00DFB407: sub esp, 00000010h
  loc_00DFB40A: mov [eax], edx
  loc_00DFB40C: mov edx, arg_14
  loc_00DFB40F: mov [eax+00000004h], ecx
  loc_00DFB412: mov ecx, var_38
  loc_00DFB415: mov [eax+00000008h], edx
  loc_00DFB418: mov edx, esp
  loc_00DFB41A: push 00000004h
  loc_00DFB41C: push FFFFFDA3h
  loc_00DFB421: mov [eax+0000000Ch], ecx
  loc_00DFB424: mov eax, var_50
  loc_00DFB427: mov ecx, arg_18
  loc_00DFB42A: mov [edx], edi
  loc_00DFB42C: push esi
  loc_00DFB42D: mov [edx+00000004h], eax
  loc_00DFB430: mov eax, var_48
  loc_00DFB433: mov [edx+00000008h], ecx
  loc_00DFB436: mov [edx+0000000Ch], eax
  loc_00DFB439: call [00401044h] ; __vbaRaiseEvent
  loc_00DFB43F: add esp, 0000004Ch
  loc_00DFB442: xor edi, edi
  loc_00DFB444: jmp 00DFB458h
  loc_00DFB446: cmp [esi+000000E2h], di
  loc_00DFB44D: jnz 00DFB458h
  loc_00DFB44F: mov ecx, [esi]
  loc_00DFB451: push esi
  loc_00DFB452: call [ecx+00000978h]
  loc_00DFB458: mov var_4, edi
  loc_00DFB45B: mov eax, Me
  loc_00DFB45E: push eax
  loc_00DFB45F: mov edx, [eax]
  loc_00DFB461: call [edx+00000008h]
  loc_00DFB464: mov eax, var_4
  loc_00DFB467: mov ecx, var_14
  loc_00DFB46A: pop edi
  loc_00DFB46B: pop esi
  loc_00DFB46C: mov fs:[00000000h], ecx
  loc_00DFB473: pop ebx
  loc_00DFB474: mov esp, ebp
  loc_00DFB476: pop ebp
  loc_00DFB477: retn 0014h
End Sub

Private Sub UserControl_UnknownEvent_14 'DFC040
  loc_00DFC040: push ebp
  loc_00DFC041: mov ebp, esp
  loc_00DFC043: sub esp, 0000000Ch
  loc_00DFC046: push 00402836h ; __vbaExceptHandler
  loc_00DFC04B: mov eax, fs:[00000000h]
  loc_00DFC051: push eax
  loc_00DFC052: mov fs:[00000000h], esp
  loc_00DFC059: sub esp, 0000001Ch
  loc_00DFC05C: push ebx
  loc_00DFC05D: push esi
  loc_00DFC05E: push edi
  loc_00DFC05F: mov var_C, esp
  loc_00DFC062: mov var_8, 00401928h
  loc_00DFC069: mov esi, Me
  loc_00DFC06C: mov eax, esi
  loc_00DFC06E: and eax, 00000001h
  loc_00DFC071: mov var_4, eax
  loc_00DFC074: and esi, FFFFFFFEh
  loc_00DFC077: push esi
  loc_00DFC078: mov Me, esi
  loc_00DFC07B: mov ecx, [esi]
  loc_00DFC07D: call [ecx+00000004h]
  loc_00DFC080: xor edx, edx
  loc_00DFC082: lea eax, var_1C
  loc_00DFC085: mov var_1C, edx
  loc_00DFC088: xor ebx, ebx
  loc_00DFC08A: push eax
  loc_00DFC08B: mov var_18, edx
  loc_00DFC08E: mov var_20, ebx
  loc_00DFC091: mov var_24, ebx
  loc_00DFC094: call 006DDC7Ch ; GetCursorPos(%x1v)
  loc_00DFC099: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DFC09F: call edi
  loc_00DFC0A1: cmp [esi+00000052h], bx
  loc_00DFC0A5: jz 00DFC0B2h
  loc_00DFC0A7: mov ecx, [esi+00000054h]
  loc_00DFC0AA: push ecx
  loc_00DFC0AB: call 006DDD70h ; SetCursor(%x1v)
  loc_00DFC0B0: call edi
  loc_00DFC0B2: mov edx, var_18
  loc_00DFC0B5: mov eax, var_1C
  loc_00DFC0B8: push edx
  loc_00DFC0B9: push eax
  loc_00DFC0BA: call 006DDC34h ; ChildWindowFromPoint(%x1v, %x2v)
  loc_00DFC0BF: mov var_20, eax
  loc_00DFC0C2: call edi
  loc_00DFC0C4: mov eax, [esi+00000010h]
  loc_00DFC0C7: lea edx, var_24
  loc_00DFC0CA: push edx
  loc_00DFC0CB: push eax
  loc_00DFC0CC: mov ecx, [eax]
  loc_00DFC0CE: call [ecx+00000058h]
  loc_00DFC0D1: cmp eax, ebx
  loc_00DFC0D3: fnclex
  loc_00DFC0D5: jge 00DFC0E9h
  loc_00DFC0D7: mov ecx, [esi+00000010h]
  loc_00DFC0DA: push 00000058h
  loc_00DFC0DC: push 006DCB00h
  loc_00DFC0E1: push ecx
  loc_00DFC0E2: push eax
  loc_00DFC0E3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC0E9: mov edx, var_20
  loc_00DFC0EC: mov eax, var_24
  loc_00DFC0EF: cmp edx, eax
  loc_00DFC0F1: jz 00DFC0F9h
  loc_00DFC0F3: mov [esi+0000004Eh], bx
  loc_00DFC0F7: jmp 00DFC13Ch
  loc_00DFC0F9: mov eax, [esi]
  loc_00DFC0FB: lea ecx, var_20
  loc_00DFC0FE: push ecx
  loc_00DFC0FF: push esi
  loc_00DFC100: mov [esi+0000004Eh], FFFFFFh
  loc_00DFC106: call [eax+00000818h]
  loc_00DFC10C: cmp eax, ebx
  loc_00DFC10E: jge 00DFC122h
  loc_00DFC110: push 00000818h
  loc_00DFC115: push 006DD090h
  loc_00DFC11A: push esi
  loc_00DFC11B: push eax
  loc_00DFC11C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC122: mov eax, var_20
  loc_00DFC125: mov edx, [esi]
  loc_00DFC127: push eax
  loc_00DFC128: push esi
  loc_00DFC129: call [edx+000009F0h]
  loc_00DFC12F: push ebx
  loc_00DFC130: push 00000001h
  loc_00DFC132: push esi
  loc_00DFC133: call [00401044h] ; __vbaRaiseEvent
  loc_00DFC139: add esp, 0000000Ch
  loc_00DFC13C: cmp [esi+000000A8h], bx
  loc_00DFC143: jnz 00DFC19Ch
  loc_00DFC145: cmp [esi+0000004Eh], bx
  loc_00DFC149: jz 00DFC18Bh
  loc_00DFC14B: cmp [esi+0000004Ch], bx
  loc_00DFC14F: jz 00DFC162h
  loc_00DFC151: mov ecx, [esi+00000048h]
  loc_00DFC154: mov eax, 00000002h
  loc_00DFC159: cmp ecx, eax
  loc_00DFC15B: jz 00DFC19Ch
  loc_00DFC15D: mov [esi+00000048h], eax
  loc_00DFC160: jmp 00DFC193h
  loc_00DFC162: mov ecx, [esi+00000048h]
  loc_00DFC165: mov eax, 00000001h
  loc_00DFC16A: cmp ecx, eax
  loc_00DFC16C: jz 00DFC19Ch
  loc_00DFC16E: mov edx, [esi]
  loc_00DFC170: push esi
  loc_00DFC171: mov [esi+00000048h], eax
  loc_00DFC174: call [edx+00000910h]
  loc_00DFC17A: cmp [esi+00000048h], 00000002h
  loc_00DFC17E: jz 00DFC19Ch
  loc_00DFC180: mov eax, [esi]
  loc_00DFC182: push esi
  loc_00DFC183: call [eax+000009B0h]
  loc_00DFC189: jmp 00DFC19Ch
  loc_00DFC18B: cmp [esi+00000048h], ebx
  loc_00DFC18E: jz 00DFC19Ch
  loc_00DFC190: mov [esi+00000048h], ebx
  loc_00DFC193: mov ecx, [esi]
  loc_00DFC195: push esi
  loc_00DFC196: call [ecx+00000910h]
  loc_00DFC19C: mov var_4, ebx
  loc_00DFC19F: mov eax, Me
  loc_00DFC1A2: push eax
  loc_00DFC1A3: mov edx, [eax]
  loc_00DFC1A5: call [edx+00000008h]
  loc_00DFC1A8: mov eax, var_4
  loc_00DFC1AB: mov ecx, var_14
  loc_00DFC1AE: pop edi
  loc_00DFC1AF: pop esi
  loc_00DFC1B0: mov fs:[00000000h], ecx
  loc_00DFC1B7: pop ebx
  loc_00DFC1B8: mov esp, ebp
  loc_00DFC1BA: pop ebp
  loc_00DFC1BB: retn 0014h
End Sub

Private Sub UserControl_UnknownEvent_15 'DFC1C0
  loc_00DFC1C0: push ebp
  loc_00DFC1C1: mov ebp, esp
  loc_00DFC1C3: sub esp, 0000000Ch
  loc_00DFC1C6: push 00402836h ; __vbaExceptHandler
  loc_00DFC1CB: mov eax, fs:[00000000h]
  loc_00DFC1D1: push eax
  loc_00DFC1D2: mov fs:[00000000h], esp
  loc_00DFC1D9: sub esp, 00000058h
  loc_00DFC1DC: push ebx
  loc_00DFC1DD: push esi
  loc_00DFC1DE: push edi
  loc_00DFC1DF: mov var_C, esp
  loc_00DFC1E2: mov var_8, 00401930h
  loc_00DFC1E9: mov esi, Me
  loc_00DFC1EC: mov eax, esi
  loc_00DFC1EE: and eax, 00000001h
  loc_00DFC1F1: mov var_4, eax
  loc_00DFC1F4: and esi, FFFFFFFEh
  loc_00DFC1F7: push esi
  loc_00DFC1F8: mov Me, esi
  loc_00DFC1FB: mov ecx, [esi]
  loc_00DFC1FD: call [ecx+00000004h]
  loc_00DFC200: xor edi, edi
  loc_00DFC202: cmp [esi+00000052h], di
  loc_00DFC206: mov var_24, edi
  loc_00DFC209: mov var_34, edi
  loc_00DFC20C: mov var_44, edi
  loc_00DFC20F: mov var_54, edi
  loc_00DFC212: mov var_58, edi
  loc_00DFC215: mov var_5C, edi
  loc_00DFC218: jz 00DFC229h
  loc_00DFC21A: mov edx, [esi+00000054h]
  loc_00DFC21D: push edx
  loc_00DFC21E: call 006DDD70h ; SetCursor(%x1v)
  loc_00DFC223: call [00401074h] ; __vbaSetSystemError
  loc_00DFC229: cmp [esi+000000E0h], di
  loc_00DFC230: jz 00DFC24Eh
  loc_00DFC232: mov eax, [esi]
  loc_00DFC234: push esi
  loc_00DFC235: mov [esi+0000004Ch], di
  loc_00DFC239: mov [esi+000000E2h], di
  loc_00DFC240: mov [esi+00000048h], edi
  loc_00DFC243: call [eax+00000910h]
  loc_00DFC249: jmp 00DFC3F0h
  loc_00DFC24E: mov ecx, Cancel
  loc_00DFC251: cmp [ecx], 0001h
  loc_00DFC255: jnz 00DFC36Bh
  loc_00DFC25B: mov edx, [esi]
  loc_00DFC25D: lea eax, var_58
  loc_00DFC260: push eax
  loc_00DFC261: push esi
  loc_00DFC262: mov [esi+0000004Ch], di
  loc_00DFC266: call [edx+00000100h]
  loc_00DFC26C: cmp eax, edi
  loc_00DFC26E: fnclex
  loc_00DFC270: jge 00DFC284h
  loc_00DFC272: push 00000100h
  loc_00DFC277: push 006DCB00h
  loc_00DFC27C: push esi
  loc_00DFC27D: push eax
  loc_00DFC27E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC284: mov ecx, [esi]
  loc_00DFC286: lea edx, var_5C
  loc_00DFC289: push edx
  loc_00DFC28A: push esi
  loc_00DFC28B: call [ecx+00000108h]
  loc_00DFC291: cmp eax, edi
  loc_00DFC293: fnclex
  loc_00DFC295: jge 00DFC2A9h
  loc_00DFC297: push 00000108h
  loc_00DFC29C: push 006DCB00h
  loc_00DFC2A1: push esi
  loc_00DFC2A2: push eax
  loc_00DFC2A3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC2A9: mov ecx, arg_14
  loc_00DFC2AC: fld real4 ptr [ecx]
  loc_00DFC2AE: fcomp real4 ptr [00401888h]
  loc_00DFC2B4: fnstsw ax
  loc_00DFC2B6: test ah, 41h
  loc_00DFC2B9: jnz 00DFC2C0h
  loc_00DFC2BB: mov edi, 00000001h
  loc_00DFC2C0: mov ebx, arg_18
  loc_00DFC2C3: fld real4 ptr [ebx]
  loc_00DFC2C5: fcomp real4 ptr [00401888h]
  loc_00DFC2CB: fnstsw ax
  loc_00DFC2CD: test ah, 41h
  loc_00DFC2D0: jnz 00DFC2D9h
  loc_00DFC2D2: mov edx, 00000001h
  loc_00DFC2D7: jmp 00DFC2DBh
  loc_00DFC2D9: xor edx, edx
  loc_00DFC2DB: fld real4 ptr [ecx]
  loc_00DFC2DD: fcomp real4 ptr var_58
  loc_00DFC2E0: fnstsw ax
  loc_00DFC2E2: test ah, 01h
  loc_00DFC2E5: jz 00DFC2EEh
  loc_00DFC2E7: mov ecx, 00000001h
  loc_00DFC2EC: jmp 00DFC2F0h
  loc_00DFC2EE: xor ecx, ecx
  loc_00DFC2F0: fld real4 ptr [ebx]
  loc_00DFC2F2: fcomp real4 ptr var_5C
  loc_00DFC2F5: fnstsw ax
  loc_00DFC2F7: test ah, 01h
  loc_00DFC2FA: jz 00DFC303h
  loc_00DFC2FC: mov eax, 00000001h
  loc_00DFC301: jmp 00DFC305h
  loc_00DFC303: xor eax, eax
  loc_00DFC305: neg eax
  loc_00DFC307: neg ecx
  loc_00DFC309: and eax, ecx
  loc_00DFC30B: neg edx
  loc_00DFC30D: and eax, edx
  loc_00DFC30F: neg edi
  loc_00DFC311: and eax, edi
  loc_00DFC313: test ax, ax
  loc_00DFC316: jz 00DFC36Bh
  loc_00DFC318: mov eax, [esi+00000064h]
  loc_00DFC31B: cmp eax, 00000001h
  loc_00DFC31E: jnz 00DFC336h
  loc_00DFC320: mov ax, [esi+0000006Ch]
  loc_00DFC324: mov ecx, [esi]
  loc_00DFC326: not ax
  loc_00DFC329: push esi
  loc_00DFC32A: mov [esi+0000006Ch], ax
  loc_00DFC32E: call [ecx+00000910h]
  loc_00DFC334: jmp 00DFC34Ah
  loc_00DFC336: cmp eax, 00000002h
  loc_00DFC339: jnz 00DFC34Ah
  loc_00DFC33B: mov edx, [esi]
  loc_00DFC33D: push esi
  loc_00DFC33E: call [edx+00000938h]
  loc_00DFC344: mov [esi+0000006Ch], FFFFFFh
  loc_00DFC34A: mov eax, [esi]
  loc_00DFC34C: push esi
  loc_00DFC34D: mov [esi+00000048h], 00000000h
  loc_00DFC354: call [eax+00000910h]
  loc_00DFC35A: push 00000000h
  loc_00DFC35C: push FFFFFDA8h
  loc_00DFC361: push esi
  loc_00DFC362: call [00401044h] ; __vbaRaiseEvent
  loc_00DFC368: add esp, 0000000Ch
  loc_00DFC36B: sub esp, 00000010h
  loc_00DFC36E: mov eax, 00004002h
  loc_00DFC373: mov ebx, esp
  loc_00DFC375: mov edx, eax
  loc_00DFC377: sub esp, 00000010h
  loc_00DFC37A: mov ecx, arg_10
  loc_00DFC37D: mov [ebx], eax
  loc_00DFC37F: mov eax, var_20
  loc_00DFC382: mov edi, 00004004h
  loc_00DFC387: mov [ebx+00000004h], eax
  loc_00DFC38A: mov eax, Cancel
  loc_00DFC38D: mov [ebx+00000008h], eax
  loc_00DFC390: mov eax, var_18
  loc_00DFC393: mov [ebx+0000000Ch], eax
  loc_00DFC396: mov eax, esp
  loc_00DFC398: sub esp, 00000010h
  loc_00DFC39B: mov [eax], edx
  loc_00DFC39D: mov edx, var_30
  loc_00DFC3A0: mov [eax+00000004h], edx
  loc_00DFC3A3: mov edx, esp
  loc_00DFC3A5: sub esp, 00000010h
  loc_00DFC3A8: mov [eax+00000008h], ecx
  loc_00DFC3AB: mov ecx, var_28
  loc_00DFC3AE: mov [eax+0000000Ch], ecx
  loc_00DFC3B1: mov eax, var_40
  loc_00DFC3B4: mov ecx, arg_14
  loc_00DFC3B7: mov [edx], edi
  loc_00DFC3B9: mov [edx+00000004h], eax
  loc_00DFC3BC: mov eax, var_38
  loc_00DFC3BF: mov [edx+00000008h], ecx
  loc_00DFC3C2: mov ecx, esp
  loc_00DFC3C4: push 00000004h
  loc_00DFC3C6: push FFFFFDA1h
  loc_00DFC3CB: mov [edx+0000000Ch], eax
  loc_00DFC3CE: mov edx, var_50
  loc_00DFC3D1: mov eax, edi
  loc_00DFC3D3: push esi
  loc_00DFC3D4: mov [ecx], eax
  loc_00DFC3D6: mov eax, arg_18
  loc_00DFC3D9: mov [ecx+00000004h], edx
  loc_00DFC3DC: mov edx, var_48
  loc_00DFC3DF: mov [ecx+00000008h], eax
  loc_00DFC3E2: mov [ecx+0000000Ch], edx
  loc_00DFC3E5: call [00401044h] ; __vbaRaiseEvent
  loc_00DFC3EB: add esp, 0000004Ch
  loc_00DFC3EE: xor edi, edi
  loc_00DFC3F0: mov var_4, edi
  loc_00DFC3F3: mov eax, Me
  loc_00DFC3F6: push eax
  loc_00DFC3F7: mov ecx, [eax]
  loc_00DFC3F9: call [ecx+00000008h]
  loc_00DFC3FC: mov eax, var_4
  loc_00DFC3FF: mov ecx, var_14
  loc_00DFC402: pop edi
  loc_00DFC403: pop esi
  loc_00DFC404: mov fs:[00000000h], ecx
  loc_00DFC40B: pop ebx
  loc_00DFC40C: mov esp, ebp
  loc_00DFC40E: pop ebp
  loc_00DFC40F: retn 0014h
End Sub

Private Sub UserControl_UnknownEvent_16 'DFC550
  loc_00DFC550: push ebp
  loc_00DFC551: mov ebp, esp
  loc_00DFC553: sub esp, 0000000Ch
  loc_00DFC556: push 00402836h ; __vbaExceptHandler
  loc_00DFC55B: mov eax, fs:[00000000h]
  loc_00DFC561: push eax
  loc_00DFC562: mov fs:[00000000h], esp
  loc_00DFC569: sub esp, 00000008h
  loc_00DFC56C: push ebx
  loc_00DFC56D: push esi
  loc_00DFC56E: push edi
  loc_00DFC56F: mov var_C, esp
  loc_00DFC572: mov var_8, 00401948h
  loc_00DFC579: mov esi, Me
  loc_00DFC57C: mov eax, esi
  loc_00DFC57E: and eax, 00000001h
  loc_00DFC581: mov var_4, eax
  loc_00DFC584: and esi, FFFFFFFEh
  loc_00DFC587: push esi
  loc_00DFC588: mov Me, esi
  loc_00DFC58B: mov ecx, [esi]
  loc_00DFC58D: call [ecx+00000004h]
  loc_00DFC590: mov edx, [esi]
  loc_00DFC592: push esi
  loc_00DFC593: call [edx+00000910h]
  loc_00DFC599: mov var_4, 00000000h
  loc_00DFC5A0: mov eax, Me
  loc_00DFC5A3: push eax
  loc_00DFC5A4: mov ecx, [eax]
  loc_00DFC5A6: call [ecx+00000008h]
  loc_00DFC5A9: mov eax, var_4
  loc_00DFC5AC: mov ecx, var_14
  loc_00DFC5AF: pop edi
  loc_00DFC5B0: pop esi
  loc_00DFC5B1: mov fs:[00000000h], ecx
  loc_00DFC5B8: pop ebx
  loc_00DFC5B9: mov esp, ebp
  loc_00DFC5BB: pop ebp
  loc_00DFC5BC: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_17 'DFA8E0
  loc_00DFA8E0: push ebp
  loc_00DFA8E1: mov ebp, esp
  loc_00DFA8E3: sub esp, 0000000Ch
  loc_00DFA8E6: push 00402836h ; __vbaExceptHandler
  loc_00DFA8EB: mov eax, fs:[00000000h]
  loc_00DFA8F1: push eax
  loc_00DFA8F2: mov fs:[00000000h], esp
  loc_00DFA8F9: sub esp, 000001CCh
  loc_00DFA8FF: push ebx
  loc_00DFA900: push esi
  loc_00DFA901: push edi
  loc_00DFA902: mov var_C, esp
  loc_00DFA905: mov var_8, 004018C8h
  loc_00DFA90C: mov esi, Me
  loc_00DFA90F: mov eax, esi
  loc_00DFA911: and eax, 00000001h
  loc_00DFA914: mov var_4, eax
  loc_00DFA917: and esi, FFFFFFFEh
  loc_00DFA91A: push esi
  loc_00DFA91B: mov Me, esi
  loc_00DFA91E: mov ecx, [esi]
  loc_00DFA920: call [ecx+00000004h]
  loc_00DFA923: mov ecx, 00000045h
  loc_00DFA928: xor eax, eax
  loc_00DFA92A: lea edi, var_12C
  loc_00DFA930: mov var_130, eax
  loc_00DFA936: repz stosd
  loc_00DFA938: mov ecx, 00000025h
  loc_00DFA93D: lea edi, var_1D0
  loc_00DFA943: repz stosd
  loc_00DFA945: mov var_134, al
  loc_00DFA94B: mov var_1D4, 00000001h
  loc_00DFA955: xor edi, edi
  loc_00DFA957: mov eax, 000000FFh
  loc_00DFA95C: cmp edi, eax
  loc_00DFA95E: jg 00DFA9E5h
  loc_00DFA964: mov ebx, [esi]
  loc_00DFA966: lea edx, var_134
  loc_00DFA96C: push edx
  loc_00DFA96D: mov ecx, edi
  loc_00DFA96F: call [00401158h] ; __vbaUI1I4
  loc_00DFA975: push eax
  loc_00DFA976: push esi
  loc_00DFA977: call [ebx+00000900h]
  loc_00DFA97D: cmp edi, 00000100h
  loc_00DFA983: jb 00DFA98Bh
  loc_00DFA985: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DFA98B: mov eax, [esi+0000018Ch]
  loc_00DFA991: mov cl, var_134
  loc_00DFA997: lea edx, var_134
  loc_00DFA99D: mov [eax+edi], cl
  loc_00DFA9A0: mov ebx, [esi]
  loc_00DFA9A2: push edx
  loc_00DFA9A3: mov ecx, edi
  loc_00DFA9A5: call [00401158h] ; __vbaUI1I4
  loc_00DFA9AB: push eax
  loc_00DFA9AC: push esi
  loc_00DFA9AD: call [ebx+00000904h]
  loc_00DFA9B3: cmp edi, 00000100h
  loc_00DFA9B9: jb 00DFA9C1h
  loc_00DFA9BB: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DFA9C1: mov edx, var_1D4
  loc_00DFA9C7: mov eax, [esi+000001A8h]
  loc_00DFA9CD: mov cl, var_134
  loc_00DFA9D3: add edx, edi
  loc_00DFA9D5: mov [eax+edi], cl
  loc_00DFA9D8: jo 00DFAAB1h
  loc_00DFA9DE: mov edi, edx
  loc_00DFA9E0: jmp 00DFA957h
  loc_00DFA9E5: lea eax, var_130
  loc_00DFA9EB: push 006DEAFCh ; "shell32.dll"
  loc_00DFA9F0: push eax
  loc_00DFA9F1: call [004011ECh] ; __vbaStrToAnsi
  loc_00DFA9F7: push eax
  loc_00DFA9F8: call 006DE4D0h ; LoadLibrary(%x1v)
  loc_00DFA9FD: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00DFAA03: mov edi, eax
  loc_00DFAA05: call ebx
  loc_00DFAA07: lea ecx, var_130
  loc_00DFAA0D: mov [esi+00000114h], edi
  loc_00DFAA13: call [00401258h] ; __vbaFreeStr
  loc_00DFAA19: lea ecx, var_12C
  loc_00DFAA1F: lea edx, var_1D0
  loc_00DFAA25: push ecx
  loc_00DFAA26: push edx
  loc_00DFAA27: push 006DD218h ; UDT_16_006DD218
  loc_00DFAA2C: mov var_12C, 00000094h
  loc_00DFAA36: call [0040113Ch] ; __vbaRecUniToAnsi
  loc_00DFAA3C: push eax
  loc_00DFAA3D: call 006DE6D8h ; GetVersionEx(%x1v)
  loc_00DFAA42: call ebx
  loc_00DFAA44: lea eax, var_1D0
  loc_00DFAA4A: lea ecx, var_12C
  loc_00DFAA50: push eax
  loc_00DFAA51: push ecx
  loc_00DFAA52: push 006DD218h ; UDT_16_006DD218
  loc_00DFAA57: call [00401058h] ; __vbaRecAnsiToUni
  loc_00DFAA5D: mov edx, var_11C
  loc_00DFAA63: xor eax, eax
  loc_00DFAA65: and edx, 00000002h
  loc_00DFAA68: cmp dl, 02h
  loc_00DFAA6B: setz al
  loc_00DFAA6E: movsx ecx, ax
  loc_00DFAA71: neg ecx
  loc_00DFAA73: mov [esi+00000078h], ecx
  loc_00DFAA76: mov var_4, 00000000h
  loc_00DFAA7D: push 00DFAA92h
  loc_00DFAA82: jmp 00DFAA91h
  loc_00DFAA84: lea ecx, var_130
  loc_00DFAA8A: call [00401258h] ; __vbaFreeStr
  loc_00DFAA90: ret
  loc_00DFAA91: ret
  loc_00DFAA92: mov eax, Me
  loc_00DFAA95: push eax
  loc_00DFAA96: mov edx, [eax]
  loc_00DFAA98: call [edx+00000008h]
  loc_00DFAA9B: mov eax, var_4
  loc_00DFAA9E: mov ecx, var_14
  loc_00DFAAA1: pop edi
  loc_00DFAAA2: pop esi
  loc_00DFAAA3: mov fs:[00000000h], ecx
  loc_00DFAAAA: pop ebx
  loc_00DFAAAB: mov esp, ebp
  loc_00DFAAAD: pop ebp
  loc_00DFAAAE: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_18 'DFEAC0
  loc_00DFEAC0: push ebp
  loc_00DFEAC1: mov ebp, esp
  loc_00DFEAC3: sub esp, 00000014h
  loc_00DFEAC6: push 00402836h ; __vbaExceptHandler
  loc_00DFEACB: mov eax, fs:[00000000h]
  loc_00DFEAD1: push eax
  loc_00DFEAD2: mov fs:[00000000h], esp
  loc_00DFEAD9: sub esp, 00000028h
  loc_00DFEADC: push ebx
  loc_00DFEADD: push esi
  loc_00DFEADE: push edi
  loc_00DFEADF: mov var_14, esp
  loc_00DFEAE2: mov var_10, 00401AC8h
  loc_00DFEAE9: mov esi, Me
  loc_00DFEAEC: mov eax, esi
  loc_00DFEAEE: and eax, 00000001h
  loc_00DFEAF1: mov var_C, eax
  loc_00DFEAF4: and esi, FFFFFFFEh
  loc_00DFEAF7: mov Me, esi
  loc_00DFEAFA: xor ebx, ebx
  loc_00DFEAFC: mov var_8, ebx
  loc_00DFEAFF: mov ecx, [esi]
  loc_00DFEB01: push esi
  loc_00DFEB02: call [ecx+00000004h]
  loc_00DFEB05: mov var_20, ebx
  loc_00DFEB08: mov var_24, ebx
  loc_00DFEB0B: push 00000001h
  loc_00DFEB0D: call [004010A4h] ; __vbaOnError
  loc_00DFEB13: mov eax, [esi+000000A4h]
  loc_00DFEB19: cmp eax, ebx
  loc_00DFEB1B: jz 00DFEB29h
  loc_00DFEB1D: push eax
  loc_00DFEB1E: call 006DD498h ; DeleteObject(%x1v)
  loc_00DFEB23: call [00401074h] ; __vbaSetSystemError
  loc_00DFEB29: mov edi, [esi]
  loc_00DFEB2B: push 006DDAD4h
  loc_00DFEB30: push ebx
  loc_00DFEB31: call [00401224h] ; __vbaCastObj
  loc_00DFEB37: push eax
  loc_00DFEB38: lea edx, var_20
  loc_00DFEB3B: push edx
  loc_00DFEB3C: call [004010ACh] ; __vbaObjSet
  loc_00DFEB42: push eax
  loc_00DFEB43: push esi
  loc_00DFEB44: call [edi+00000A34h]
  loc_00DFEB4A: fnclex
  loc_00DFEB4C: cmp eax, ebx
  loc_00DFEB4E: jge 00DFEB66h
  loc_00DFEB50: push 00000A34h
  loc_00DFEB55: push 006DD090h
  loc_00DFEB5A: push esi
  loc_00DFEB5B: push eax
  loc_00DFEB5C: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEB62: call edi
  loc_00DFEB64: jmp 00DFEB6Ch
  loc_00DFEB66: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEB6C: lea ecx, var_20
  loc_00DFEB6F: call [00401254h] ; __vbaFreeObj
  loc_00DFEB75: mov eax, [esi+00000114h]
  loc_00DFEB7B: push eax
  loc_00DFEB7C: call 006DE488h ; FreeLibrary(%x1v)
  loc_00DFEB81: call [00401074h] ; __vbaSetSystemError
  loc_00DFEB87: mov ecx, [esi]
  loc_00DFEB89: push esi
  loc_00DFEB8A: call [ecx+000007ACh]
  loc_00DFEB90: cmp eax, ebx
  loc_00DFEB92: jge 00DFEBA2h
  loc_00DFEB94: push 000007ACh
  loc_00DFEB99: push 006DD090h
  loc_00DFEB9E: push esi
  loc_00DFEB9F: push eax
  loc_00DFEBA0: call edi
  loc_00DFEBA2: mov edx, [esi]
  loc_00DFEBA4: lea eax, var_20
  loc_00DFEBA7: push eax
  loc_00DFEBA8: push esi
  loc_00DFEBA9: call [edx+000002B0h]
  loc_00DFEBAF: fnclex
  loc_00DFEBB1: cmp eax, ebx
  loc_00DFEBB3: jge 00DFEBC3h
  loc_00DFEBB5: push 000002B0h
  loc_00DFEBBA: push 006DCB00h
  loc_00DFEBBF: push esi
  loc_00DFEBC0: push eax
  loc_00DFEBC1: call edi
  loc_00DFEBC3: mov eax, var_20
  loc_00DFEBC6: mov ebx, eax
  loc_00DFEBC8: mov ecx, [eax]
  loc_00DFEBCA: lea edx, var_24
  loc_00DFEBCD: push edx
  loc_00DFEBCE: push eax
  loc_00DFEBCF: call [ecx+0000003Ch]
  loc_00DFEBD2: fnclex
  loc_00DFEBD4: test eax, eax
  loc_00DFEBD6: jge 00DFEBE3h
  loc_00DFEBD8: push 0000003Ch
  loc_00DFEBDA: push 006DEA70h
  loc_00DFEBDF: push ebx
  loc_00DFEBE0: push eax
  loc_00DFEBE1: call edi
  loc_00DFEBE3: mov edi, var_24
  loc_00DFEBE6: lea ecx, var_20
  loc_00DFEBE9: call [00401254h] ; __vbaFreeObj
  loc_00DFEBEF: test di, di
  loc_00DFEBF2: jz 00DFEC06h
  loc_00DFEBF4: mov eax, [esi]
  loc_00DFEBF6: push esi
  loc_00DFEBF7: call [eax+00000A28h]
  loc_00DFEBFD: mov ecx, [esi]
  loc_00DFEBFF: push esi
  loc_00DFEC00: call [ecx+00000A28h]
  loc_00DFEC06: call [00401098h] ; __vbaExitProc
  loc_00DFEC0C: push 00DFEC1Eh
  loc_00DFEC11: jmp 00DFEC1Dh
  loc_00DFEC13: lea ecx, var_20
  loc_00DFEC16: call [00401254h] ; __vbaFreeObj
  loc_00DFEC1C: ret
  loc_00DFEC1D: ret
  loc_00DFEC1E: mov eax, Me
  loc_00DFEC21: mov edx, [eax]
  loc_00DFEC23: push eax
  loc_00DFEC24: call [edx+00000008h]
  loc_00DFEC27: mov eax, var_C
  loc_00DFEC2A: mov ecx, var_1C
  loc_00DFEC2D: mov fs:[00000000h], ecx
  loc_00DFEC34: pop edi
  loc_00DFEC35: pop esi
  loc_00DFEC36: pop ebx
  loc_00DFEC37: mov esp, ebp
  loc_00DFEC39: pop ebp
  loc_00DFEC3A: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_1F 'DFEC40
  loc_00DFEC40: push ebp
  loc_00DFEC41: mov ebp, esp
  loc_00DFEC43: sub esp, 0000000Ch
  loc_00DFEC46: push 00402836h ; __vbaExceptHandler
  loc_00DFEC4B: mov eax, fs:[00000000h]
  loc_00DFEC51: push eax
  loc_00DFEC52: mov fs:[00000000h], esp
  loc_00DFEC59: sub esp, 0000008Ch
  loc_00DFEC5F: push ebx
  loc_00DFEC60: push esi
  loc_00DFEC61: push edi
  loc_00DFEC62: mov var_C, esp
  loc_00DFEC65: mov var_8, 00401AF0h
  loc_00DFEC6C: mov esi, Me
  loc_00DFEC6F: mov eax, esi
  loc_00DFEC71: and eax, 00000001h
  loc_00DFEC74: mov var_4, eax
  loc_00DFEC77: and esi, FFFFFFFEh
  loc_00DFEC7A: push esi
  loc_00DFEC7B: mov Me, esi
  loc_00DFEC7E: mov ecx, [esi]
  loc_00DFEC80: call [ecx+00000004h]
  loc_00DFEC83: mov edx, Cancel
  loc_00DFEC86: xor eax, eax
  loc_00DFEC88: mov var_18, eax
  loc_00DFEC8B: mov var_1C, eax
  loc_00DFEC8E: mov var_20, eax
  loc_00DFEC91: mov var_30, eax
  loc_00DFEC94: mov var_40, eax
  loc_00DFEC97: mov var_50, eax
  loc_00DFEC9A: mov var_64, eax
  loc_00DFEC9D: mov var_68, eax
  loc_00DFECA0: mov var_6C, eax
  loc_00DFECA3: mov var_84, eax
  loc_00DFECA9: mov eax, [edx]
  loc_00DFECAB: lea ecx, var_84
  loc_00DFECB1: push eax
  loc_00DFECB2: push ecx
  loc_00DFECB3: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFECB9: mov edx, [esi+00000044h]
  loc_00DFECBC: mov edi, var_84
  loc_00DFECC2: mov ecx, 00000003h
  loc_00DFECC7: mov var_48, edx
  loc_00DFECCA: mov edx, ecx
  loc_00DFECCC: sub esp, 00000010h
  loc_00DFECCF: mov var_50, edx
  loc_00DFECD2: mov ebx, [edi]
  loc_00DFECD4: mov edi, esp
  loc_00DFECD6: mov eax, 00000001h
  loc_00DFECDB: mov var_A0, ebx
  loc_00DFECE1: mov ebx, var_54
  loc_00DFECE4: mov [edi], ecx
  loc_00DFECE6: mov ecx, esp
  loc_00DFECE8: mov var_9C, edi
  loc_00DFECEE: mov edi, var_5C
  loc_00DFECF1: mov [ecx+00000004h], edi
  loc_00DFECF4: sub esp, 00000010h
  loc_00DFECF7: mov [ecx+00000008h], eax
  loc_00DFECFA: mov eax, esp
  loc_00DFECFC: push 006DEB60h ; "ButtonStyle"
  loc_00DFED01: mov [ecx+0000000Ch], ebx
  loc_00DFED04: mov ecx, var_4C
  loc_00DFED07: mov [eax], edx
  loc_00DFED09: mov edx, var_48
  loc_00DFED0C: mov [eax+00000004h], ecx
  loc_00DFED0F: mov ecx, var_44
  loc_00DFED12: mov [eax+00000008h], edx
  loc_00DFED15: mov edx, var_84
  loc_00DFED1B: push edx
  loc_00DFED1C: mov [eax+0000000Ch], ecx
  loc_00DFED1F: mov eax, var_A0
  loc_00DFED25: call [eax+00000020h]
  loc_00DFED28: test eax, eax
  loc_00DFED2A: fnclex
  loc_00DFED2C: jge 00DFED43h
  loc_00DFED2E: mov ecx, var_84
  loc_00DFED34: push 00000020h
  loc_00DFED36: push 006DEB78h
  loc_00DFED3B: push ecx
  loc_00DFED3C: push eax
  loc_00DFED3D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFED43: mov dx, [esi+0000006Eh]
  loc_00DFED47: mov ecx, var_84
  loc_00DFED4D: mov eax, 0000000Bh
  loc_00DFED52: mov var_48, dx
  loc_00DFED56: mov var_50, eax
  loc_00DFED59: mov edx, [ecx]
  loc_00DFED5B: sub esp, 00000010h
  loc_00DFED5E: mov ecx, esp
  loc_00DFED60: sub esp, 00000010h
  loc_00DFED63: mov [ecx], eax
  loc_00DFED65: xor eax, eax
  loc_00DFED67: mov [ecx+00000004h], edi
  loc_00DFED6A: mov [ecx+00000008h], eax
  loc_00DFED6D: mov eax, esp
  loc_00DFED6F: push 006DEB8Ch ; "ShowFocusRect"
  loc_00DFED74: mov [ecx+0000000Ch], ebx
  loc_00DFED77: mov ecx, var_50
  loc_00DFED7A: mov [eax], ecx
  loc_00DFED7C: mov ecx, var_4C
  loc_00DFED7F: mov [eax+00000004h], ecx
  loc_00DFED82: mov ecx, var_48
  loc_00DFED85: mov [eax+00000008h], ecx
  loc_00DFED88: mov ecx, var_44
  loc_00DFED8B: mov [eax+0000000Ch], ecx
  loc_00DFED8E: mov eax, var_84
  loc_00DFED94: push eax
  loc_00DFED95: call [edx+00000020h]
  loc_00DFED98: test eax, eax
  loc_00DFED9A: fnclex
  loc_00DFED9C: jge 00DFEDB3h
  loc_00DFED9E: mov ecx, var_84
  loc_00DFEDA4: push 00000020h
  loc_00DFEDA6: push 006DEB78h
  loc_00DFEDAB: push ecx
  loc_00DFEDAC: push eax
  loc_00DFEDAD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEDB3: lea ecx, var_50
  loc_00DFEDB6: call [00401024h] ; __vbaFreeVar
  loc_00DFEDBC: mov dx, [esi+0000007Ch]
  loc_00DFEDC0: mov ecx, var_84
  loc_00DFEDC6: mov eax, 0000000Bh
  loc_00DFEDCB: mov var_48, dx
  loc_00DFEDCF: mov var_50, eax
  loc_00DFEDD2: mov edx, [ecx]
  loc_00DFEDD4: sub esp, 00000010h
  loc_00DFEDD7: mov ecx, esp
  loc_00DFEDD9: sub esp, 00000010h
  loc_00DFEDDC: mov [ecx], eax
  loc_00DFEDDE: or eax, FFFFFFFFh
  loc_00DFEDE1: mov [ecx+00000004h], edi
  loc_00DFEDE4: mov [ecx+00000008h], eax
  loc_00DFEDE7: mov eax, esp
  loc_00DFEDE9: push 006DEBBCh ; "Enabled"
  loc_00DFEDEE: mov [ecx+0000000Ch], ebx
  loc_00DFEDF1: mov ecx, var_50
  loc_00DFEDF4: mov [eax], ecx
  loc_00DFEDF6: mov ecx, var_4C
  loc_00DFEDF9: mov [eax+00000004h], ecx
  loc_00DFEDFC: mov ecx, var_48
  loc_00DFEDFF: mov [eax+00000008h], ecx
  loc_00DFEE02: mov ecx, var_44
  loc_00DFEE05: mov [eax+0000000Ch], ecx
  loc_00DFEE08: mov eax, var_84
  loc_00DFEE0E: push eax
  loc_00DFEE0F: call [edx+00000020h]
  loc_00DFEE12: test eax, eax
  loc_00DFEE14: fnclex
  loc_00DFEE16: jge 00DFEE2Dh
  loc_00DFEE18: mov ecx, var_84
  loc_00DFEE1E: push 00000020h
  loc_00DFEE20: push 006DEB78h
  loc_00DFEE25: push ecx
  loc_00DFEE26: push eax
  loc_00DFEE27: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEE2D: lea ecx, var_50
  loc_00DFEE30: call [00401024h] ; __vbaFreeVar
  loc_00DFEE36: mov edx, [esi]
  loc_00DFEE38: lea eax, var_18
  loc_00DFEE3B: push eax
  loc_00DFEE3C: push esi
  loc_00DFEE3D: call [edx+00000A2Ch]
  loc_00DFEE43: test eax, eax
  loc_00DFEE45: fnclex
  loc_00DFEE47: jge 00DFEE5Bh
  loc_00DFEE49: push 00000A2Ch
  loc_00DFEE4E: push 006DD090h
  loc_00DFEE53: push esi
  loc_00DFEE54: push eax
  loc_00DFEE55: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEE5B: mov ecx, [esi]
  loc_00DFEE5D: lea edx, var_1C
  loc_00DFEE60: push edx
  loc_00DFEE61: push esi
  loc_00DFEE62: call [ecx+000002B0h]
  loc_00DFEE68: test eax, eax
  loc_00DFEE6A: fnclex
  loc_00DFEE6C: jge 00DFEE80h
  loc_00DFEE6E: push 000002B0h
  loc_00DFEE73: push 006DCB00h
  loc_00DFEE78: push esi
  loc_00DFEE79: push eax
  loc_00DFEE7A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEE80: mov eax, var_1C
  loc_00DFEE83: lea edx, var_20
  loc_00DFEE86: push edx
  loc_00DFEE87: push eax
  loc_00DFEE88: mov ecx, [eax]
  loc_00DFEE8A: mov var_78, eax
  loc_00DFEE8D: call [ecx+00000024h]
  loc_00DFEE90: test eax, eax
  loc_00DFEE92: fnclex
  loc_00DFEE94: jge 00DFEEA8h
  loc_00DFEE96: mov ecx, var_78
  loc_00DFEE99: push 00000024h
  loc_00DFEE9B: push 006DEA70h
  loc_00DFEEA0: push ecx
  loc_00DFEEA1: push eax
  loc_00DFEEA2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEEA8: mov eax, var_20
  loc_00DFEEAB: mov ecx, var_18
  loc_00DFEEAE: mov edx, var_84
  loc_00DFEEB4: mov var_38, eax
  loc_00DFEEB7: mov eax, 00000009h
  loc_00DFEEBC: mov var_20, 00000000h
  loc_00DFEEC3: mov var_40, eax
  loc_00DFEEC6: mov var_18, 00000000h
  loc_00DFEECD: mov var_28, ecx
  loc_00DFEED0: mov var_30, eax
  loc_00DFEED3: mov ecx, [edx]
  loc_00DFEED5: sub esp, 00000010h
  loc_00DFEED8: mov edx, esp
  loc_00DFEEDA: sub esp, 00000010h
  loc_00DFEEDD: mov [edx], eax
  loc_00DFEEDF: mov eax, var_3C
  loc_00DFEEE2: mov [edx+00000004h], eax
  loc_00DFEEE5: mov eax, var_38
  loc_00DFEEE8: mov [edx+00000008h], eax
  loc_00DFEEEB: mov eax, var_34
  loc_00DFEEEE: mov [edx+0000000Ch], eax
  loc_00DFEEF1: mov eax, var_30
  loc_00DFEEF4: mov edx, esp
  loc_00DFEEF6: push 006DEBACh ; "Font"
  loc_00DFEEFB: mov [edx], eax
  loc_00DFEEFD: mov eax, var_2C
  loc_00DFEF00: mov [edx+00000004h], eax
  loc_00DFEF03: mov eax, var_28
  loc_00DFEF06: mov [edx+00000008h], eax
  loc_00DFEF09: mov eax, var_24
  loc_00DFEF0C: mov [edx+0000000Ch], eax
  loc_00DFEF0F: mov edx, var_84
  loc_00DFEF15: push edx
  loc_00DFEF16: call [ecx+00000020h]
  loc_00DFEF19: test eax, eax
  loc_00DFEF1B: fnclex
  loc_00DFEF1D: jge 00DFEF34h
  loc_00DFEF1F: mov ecx, var_84
  loc_00DFEF25: push 00000020h
  loc_00DFEF27: push 006DEB78h
  loc_00DFEF2C: push ecx
  loc_00DFEF2D: push eax
  loc_00DFEF2E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEF34: lea ecx, var_1C
  loc_00DFEF37: call [00401254h] ; __vbaFreeObj
  loc_00DFEF3D: lea edx, var_40
  loc_00DFEF40: lea eax, var_30
  loc_00DFEF43: push edx
  loc_00DFEF44: push eax
  loc_00DFEF45: push 00000002h
  loc_00DFEF47: call [00401038h] ; __vbaFreeVarList
  loc_00DFEF4D: add esp, 0000000Ch
  loc_00DFEF50: push 0000000Fh
  loc_00DFEF52: call 006DDDB4h ; GetSysColor(%x1v)
  loc_00DFEF57: mov var_68, eax
  loc_00DFEF5A: call [00401074h] ; __vbaSetSystemError
  loc_00DFEF60: mov ecx, [esi+00000088h]
  loc_00DFEF66: mov edx, var_84
  loc_00DFEF6C: mov eax, 00000003h
  loc_00DFEF71: mov var_48, ecx
  loc_00DFEF74: mov var_50, eax
  loc_00DFEF77: mov ecx, [edx]
  loc_00DFEF79: sub esp, 00000010h
  loc_00DFEF7C: mov edx, esp
  loc_00DFEF7E: sub esp, 00000010h
  loc_00DFEF81: mov [edx], eax
  loc_00DFEF83: mov eax, var_68
  loc_00DFEF86: mov [edx+00000004h], edi
  loc_00DFEF89: mov [edx+00000008h], eax
  loc_00DFEF8C: mov eax, esp
  loc_00DFEF8E: push 006DEACCh ; "BackColor"
  loc_00DFEF93: mov [edx+0000000Ch], ebx
  loc_00DFEF96: mov edx, var_50
  loc_00DFEF99: mov [eax], edx
  loc_00DFEF9B: mov edx, var_4C
  loc_00DFEF9E: mov [eax+00000004h], edx
  loc_00DFEFA1: mov edx, var_48
  loc_00DFEFA4: mov [eax+00000008h], edx
  loc_00DFEFA7: mov edx, var_44
  loc_00DFEFAA: mov [eax+0000000Ch], edx
  loc_00DFEFAD: mov eax, var_84
  loc_00DFEFB3: push eax
  loc_00DFEFB4: call [ecx+00000020h]
  loc_00DFEFB7: test eax, eax
  loc_00DFEFB9: fnclex
  loc_00DFEFBB: jge 00DFEFD2h
  loc_00DFEFBD: mov ecx, var_84
  loc_00DFEFC3: push 00000020h
  loc_00DFEFC5: push 006DEB78h
  loc_00DFEFCA: push ecx
  loc_00DFEFCB: push eax
  loc_00DFEFCC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEFD2: mov edx, [esi+00000080h]
  loc_00DFEFD8: mov ecx, var_84
  loc_00DFEFDE: mov eax, 00000008h
  loc_00DFEFE3: mov var_48, edx
  loc_00DFEFE6: mov var_50, eax
  loc_00DFEFE9: mov edx, [ecx]
  loc_00DFEFEB: sub esp, 00000010h
  loc_00DFEFEE: mov ecx, esp
  loc_00DFEFF0: sub esp, 00000010h
  loc_00DFEFF3: mov [ecx], eax
  loc_00DFEFF5: mov eax, 006DF088h ; "jcbutton1"
  loc_00DFEFFA: mov [ecx+00000004h], edi
  loc_00DFEFFD: mov [ecx+00000008h], eax
  loc_00DFF000: mov eax, esp
  loc_00DFF002: push 006DEBD0h ; "Caption"
  loc_00DFF007: mov [ecx+0000000Ch], ebx
  loc_00DFF00A: mov ecx, var_50
  loc_00DFF00D: mov [eax], ecx
  loc_00DFF00F: mov ecx, var_4C
  loc_00DFF012: mov [eax+00000004h], ecx
  loc_00DFF015: mov ecx, var_48
  loc_00DFF018: mov [eax+00000008h], ecx
  loc_00DFF01B: mov ecx, var_44
  loc_00DFF01E: mov [eax+0000000Ch], ecx
  loc_00DFF021: mov eax, var_84
  loc_00DFF027: push eax
  loc_00DFF028: call [edx+00000020h]
  loc_00DFF02B: test eax, eax
  loc_00DFF02D: fnclex
  loc_00DFF02F: jge 00DFF046h
  loc_00DFF031: mov ecx, var_84
  loc_00DFF037: push 00000020h
  loc_00DFF039: push 006DEB78h
  loc_00DFF03E: push ecx
  loc_00DFF03F: push eax
  loc_00DFF040: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF046: mov edx, [esi]
  loc_00DFF048: lea eax, var_6C
  loc_00DFF04B: lea ecx, var_68
  loc_00DFF04E: push eax
  loc_00DFF04F: push ecx
  loc_00DFF050: push 80000012h
  loc_00DFF055: push esi
  loc_00DFF056: mov var_68, 00000000h
  loc_00DFF05D: call [edx+0000090Ch]
  loc_00DFF063: mov edx, [esi+00000090h]
  loc_00DFF069: mov ecx, var_84
  loc_00DFF06F: mov eax, 00000003h
  loc_00DFF074: mov var_48, edx
  loc_00DFF077: mov var_50, eax
  loc_00DFF07A: mov edx, [ecx]
  loc_00DFF07C: sub esp, 00000010h
  loc_00DFF07F: mov ecx, esp
  loc_00DFF081: sub esp, 00000010h
  loc_00DFF084: mov [ecx], eax
  loc_00DFF086: mov eax, var_6C
  loc_00DFF089: mov [ecx+00000004h], edi
  loc_00DFF08C: mov [ecx+00000008h], eax
  loc_00DFF08F: mov eax, esp
  loc_00DFF091: push 006DEE94h ; "ForeColor"
  loc_00DFF096: mov [ecx+0000000Ch], ebx
  loc_00DFF099: mov ecx, var_50
  loc_00DFF09C: mov [eax], ecx
  loc_00DFF09E: mov ecx, var_4C
  loc_00DFF0A1: mov [eax+00000004h], ecx
  loc_00DFF0A4: mov ecx, var_48
  loc_00DFF0A7: mov [eax+00000008h], ecx
  loc_00DFF0AA: mov ecx, var_44
  loc_00DFF0AD: mov [eax+0000000Ch], ecx
  loc_00DFF0B0: mov eax, var_84
  loc_00DFF0B6: push eax
  loc_00DFF0B7: call [edx+00000020h]
  loc_00DFF0BA: test eax, eax
  loc_00DFF0BC: fnclex
  loc_00DFF0BE: jge 00DFF0D5h
  loc_00DFF0C0: mov ecx, var_84
  loc_00DFF0C6: push 00000020h
  loc_00DFF0C8: push 006DEB78h
  loc_00DFF0CD: push ecx
  loc_00DFF0CE: push eax
  loc_00DFF0CF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF0D5: mov edx, [esi]
  loc_00DFF0D7: lea eax, var_6C
  loc_00DFF0DA: lea ecx, var_68
  loc_00DFF0DD: push eax
  loc_00DFF0DE: push ecx
  loc_00DFF0DF: push 80000012h
  loc_00DFF0E4: push esi
  loc_00DFF0E5: mov var_68, 00000000h
  loc_00DFF0EC: call [edx+0000090Ch]
  loc_00DFF0F2: mov edx, [esi+00000094h]
  loc_00DFF0F8: mov ecx, var_84
  loc_00DFF0FE: mov eax, 00000003h
  loc_00DFF103: mov var_48, edx
  loc_00DFF106: mov var_50, eax
  loc_00DFF109: mov edx, [ecx]
  loc_00DFF10B: sub esp, 00000010h
  loc_00DFF10E: mov ecx, esp
  loc_00DFF110: sub esp, 00000010h
  loc_00DFF113: mov [ecx], eax
  loc_00DFF115: mov eax, var_6C
  loc_00DFF118: mov [ecx+00000004h], edi
  loc_00DFF11B: mov [ecx+00000008h], eax
  loc_00DFF11E: mov eax, esp
  loc_00DFF120: push 006DEEACh ; "ForeColorHover"
  loc_00DFF125: mov [ecx+0000000Ch], ebx
  loc_00DFF128: mov ecx, var_50
  loc_00DFF12B: mov [eax], ecx
  loc_00DFF12D: mov ecx, var_4C
  loc_00DFF130: mov [eax+00000004h], ecx
  loc_00DFF133: mov ecx, var_48
  loc_00DFF136: mov [eax+00000008h], ecx
  loc_00DFF139: mov ecx, var_44
  loc_00DFF13C: mov [eax+0000000Ch], ecx
  loc_00DFF13F: mov eax, var_84
  loc_00DFF145: push eax
  loc_00DFF146: call [edx+00000020h]
  loc_00DFF149: test eax, eax
  loc_00DFF14B: fnclex
  loc_00DFF14D: jge 00DFF164h
  loc_00DFF14F: mov ecx, var_84
  loc_00DFF155: push 00000020h
  loc_00DFF157: push 006DEB78h
  loc_00DFF15C: push ecx
  loc_00DFF15D: push eax
  loc_00DFF15E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF164: mov edx, [esi+00000064h]
  loc_00DFF167: mov ecx, var_84
  loc_00DFF16D: mov eax, 00000003h
  loc_00DFF172: mov var_48, edx
  loc_00DFF175: mov var_50, eax
  loc_00DFF178: mov edx, [ecx]
  loc_00DFF17A: sub esp, 00000010h
  loc_00DFF17D: mov ecx, esp
  loc_00DFF17F: sub esp, 00000010h
  loc_00DFF182: mov [ecx], eax
  loc_00DFF184: xor eax, eax
  loc_00DFF186: mov [ecx+00000004h], edi
  loc_00DFF189: mov [ecx+00000008h], eax
  loc_00DFF18C: mov eax, esp
  loc_00DFF18E: push 006DEE44h ; "Mode"
  loc_00DFF193: mov [ecx+0000000Ch], ebx
  loc_00DFF196: mov ecx, var_50
  loc_00DFF199: mov [eax], ecx
  loc_00DFF19B: mov ecx, var_4C
  loc_00DFF19E: mov [eax+00000004h], ecx
  loc_00DFF1A1: mov ecx, var_48
  loc_00DFF1A4: mov [eax+00000008h], ecx
  loc_00DFF1A7: mov ecx, var_44
  loc_00DFF1AA: mov [eax+0000000Ch], ecx
  loc_00DFF1AD: mov eax, var_84
  loc_00DFF1B3: push eax
  loc_00DFF1B4: call [edx+00000020h]
  loc_00DFF1B7: test eax, eax
  loc_00DFF1B9: fnclex
  loc_00DFF1BB: jge 00DFF1D2h
  loc_00DFF1BD: mov ecx, var_84
  loc_00DFF1C3: push 00000020h
  loc_00DFF1C5: push 006DEB78h
  loc_00DFF1CA: push ecx
  loc_00DFF1CB: push eax
  loc_00DFF1CC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF1D2: mov dx, [esi+0000006Ch]
  loc_00DFF1D6: mov ecx, var_84
  loc_00DFF1DC: mov eax, 0000000Bh
  loc_00DFF1E1: mov var_48, dx
  loc_00DFF1E5: mov var_50, eax
  loc_00DFF1E8: mov edx, [ecx]
  loc_00DFF1EA: sub esp, 00000010h
  loc_00DFF1ED: mov ecx, esp
  loc_00DFF1EF: sub esp, 00000010h
  loc_00DFF1F2: mov [ecx], eax
  loc_00DFF1F4: xor eax, eax
  loc_00DFF1F6: mov [ecx+00000004h], edi
  loc_00DFF1F9: mov [ecx+00000008h], eax
  loc_00DFF1FC: mov eax, esp
  loc_00DFF1FE: push 006DEBFCh ; "Value"
  loc_00DFF203: mov [ecx+0000000Ch], ebx
  loc_00DFF206: mov ecx, var_50
  loc_00DFF209: mov [eax], ecx
  loc_00DFF20B: mov ecx, var_4C
  loc_00DFF20E: mov [eax+00000004h], ecx
  loc_00DFF211: mov ecx, var_48
  loc_00DFF214: mov [eax+00000008h], ecx
  loc_00DFF217: mov ecx, var_44
  loc_00DFF21A: mov [eax+0000000Ch], ecx
  loc_00DFF21D: mov eax, var_84
  loc_00DFF223: push eax
  loc_00DFF224: call [edx+00000020h]
  loc_00DFF227: test eax, eax
  loc_00DFF229: fnclex
  loc_00DFF22B: jge 00DFF242h
  loc_00DFF22D: mov ecx, var_84
  loc_00DFF233: push 00000020h
  loc_00DFF235: push 006DEB78h
  loc_00DFF23A: push ecx
  loc_00DFF23B: push eax
  loc_00DFF23C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF242: lea ecx, var_50
  loc_00DFF245: call [00401024h] ; __vbaFreeVar
  loc_00DFF24B: mov eax, [esi+00000010h]
  loc_00DFF24E: lea ecx, var_64
  loc_00DFF251: push ecx
  loc_00DFF252: push eax
  loc_00DFF253: mov edx, [eax]
  loc_00DFF255: call [edx+000000A0h]
  loc_00DFF25B: test eax, eax
  loc_00DFF25D: fnclex
  loc_00DFF25F: jge 00DFF276h
  loc_00DFF261: mov edx, [esi+00000010h]
  loc_00DFF264: push 000000A0h
  loc_00DFF269: push 006DCB00h
  loc_00DFF26E: push edx
  loc_00DFF26F: push eax
  loc_00DFF270: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF276: mov cx, var_64
  loc_00DFF27A: mov edx, var_84
  loc_00DFF280: mov eax, 00000002h
  loc_00DFF285: mov var_48, cx
  loc_00DFF289: mov var_50, eax
  loc_00DFF28C: mov ecx, [edx]
  loc_00DFF28E: sub esp, 00000010h
  loc_00DFF291: mov edx, esp
  loc_00DFF293: sub esp, 00000010h
  loc_00DFF296: mov [edx], eax
  loc_00DFF298: xor eax, eax
  loc_00DFF29A: mov [edx+00000004h], edi
  loc_00DFF29D: mov [edx+00000008h], eax
  loc_00DFF2A0: mov eax, esp
  loc_00DFF2A2: push 006DEC0Ch ; "MousePointer"
  loc_00DFF2A7: mov [edx+0000000Ch], ebx
  loc_00DFF2AA: mov edx, var_50
  loc_00DFF2AD: mov [eax], edx
  loc_00DFF2AF: mov edx, var_4C
  loc_00DFF2B2: mov [eax+00000004h], edx
  loc_00DFF2B5: mov edx, var_48
  loc_00DFF2B8: mov [eax+00000008h], edx
  loc_00DFF2BB: mov edx, var_44
  loc_00DFF2BE: mov [eax+0000000Ch], edx
  loc_00DFF2C1: mov eax, var_84
  loc_00DFF2C7: push eax
  loc_00DFF2C8: call [ecx+00000020h]
  loc_00DFF2CB: test eax, eax
  loc_00DFF2CD: fnclex
  loc_00DFF2CF: jge 00DFF2E6h
  loc_00DFF2D1: mov ecx, var_84
  loc_00DFF2D7: push 00000020h
  loc_00DFF2D9: push 006DEB78h
  loc_00DFF2DE: push ecx
  loc_00DFF2DF: push eax
  loc_00DFF2E0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF2E6: mov dx, [esi+00000052h]
  loc_00DFF2EA: mov ecx, var_84
  loc_00DFF2F0: mov eax, 0000000Bh
  loc_00DFF2F5: mov var_48, dx
  loc_00DFF2F9: mov var_50, eax
  loc_00DFF2FC: mov edx, [ecx]
  loc_00DFF2FE: sub esp, 00000010h
  loc_00DFF301: mov ecx, esp
  loc_00DFF303: sub esp, 00000010h
  loc_00DFF306: mov [ecx], eax
  loc_00DFF308: xor eax, eax
  loc_00DFF30A: mov [ecx+00000004h], edi
  loc_00DFF30D: mov [ecx+00000008h], eax
  loc_00DFF310: mov eax, esp
  loc_00DFF312: push 006DEC2Ch ; "HandPointer"
  loc_00DFF317: mov [ecx+0000000Ch], ebx
  loc_00DFF31A: mov ecx, var_50
  loc_00DFF31D: mov [eax], ecx
  loc_00DFF31F: mov ecx, var_4C
  loc_00DFF322: mov [eax+00000004h], ecx
  loc_00DFF325: mov ecx, var_48
  loc_00DFF328: mov [eax+00000008h], ecx
  loc_00DFF32B: mov ecx, var_44
  loc_00DFF32E: mov [eax+0000000Ch], ecx
  loc_00DFF331: mov eax, var_84
  loc_00DFF337: push eax
  loc_00DFF338: call [edx+00000020h]
  loc_00DFF33B: test eax, eax
  loc_00DFF33D: fnclex
  loc_00DFF33F: jge 00DFF356h
  loc_00DFF341: mov ecx, var_84
  loc_00DFF347: push 00000020h
  loc_00DFF349: push 006DEB78h
  loc_00DFF34E: push ecx
  loc_00DFF34F: push eax
  loc_00DFF350: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF356: lea ecx, var_50
  loc_00DFF359: call [00401024h] ; __vbaFreeVar
  loc_00DFF35F: mov eax, [esi+00000010h]
  loc_00DFF362: lea ecx, var_18
  loc_00DFF365: push ecx
  loc_00DFF366: push eax
  loc_00DFF367: mov edx, [eax]
  loc_00DFF369: call [edx+00000220h]
  loc_00DFF36F: test eax, eax
  loc_00DFF371: fnclex
  loc_00DFF373: jge 00DFF38Ah
  loc_00DFF375: mov edx, [esi+00000010h]
  loc_00DFF378: push 00000220h
  loc_00DFF37D: push 006DCB00h
  loc_00DFF382: push edx
  loc_00DFF383: push eax
  loc_00DFF384: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF38A: mov ecx, var_18
  loc_00DFF38D: mov eax, 00000009h
  loc_00DFF392: mov var_28, ecx
  loc_00DFF395: mov ecx, var_84
  loc_00DFF39B: mov var_38, 00000000h
  loc_00DFF3A2: mov var_40, eax
  loc_00DFF3A5: mov var_18, 00000000h
  loc_00DFF3AC: mov var_30, eax
  loc_00DFF3AF: mov edx, [ecx]
  loc_00DFF3B1: sub esp, 00000010h
  loc_00DFF3B4: mov ecx, esp
  loc_00DFF3B6: sub esp, 00000010h
  loc_00DFF3B9: mov [ecx], eax
  loc_00DFF3BB: mov eax, var_3C
  loc_00DFF3BE: mov [ecx+00000004h], eax
  loc_00DFF3C1: mov eax, var_38
  loc_00DFF3C4: mov [ecx+00000008h], eax
  loc_00DFF3C7: mov eax, var_34
  loc_00DFF3CA: mov [ecx+0000000Ch], eax
  loc_00DFF3CD: mov eax, var_30
  loc_00DFF3D0: mov ecx, esp
  loc_00DFF3D2: push 006DEC54h ; "MouseIcon"
  loc_00DFF3D7: mov [ecx], eax
  loc_00DFF3D9: mov eax, var_2C
  loc_00DFF3DC: mov [ecx+00000004h], eax
  loc_00DFF3DF: mov eax, var_28
  loc_00DFF3E2: mov [ecx+00000008h], eax
  loc_00DFF3E5: mov eax, var_24
  loc_00DFF3E8: mov [ecx+0000000Ch], eax
  loc_00DFF3EB: mov ecx, var_84
  loc_00DFF3F1: push ecx
  loc_00DFF3F2: call [edx+00000020h]
  loc_00DFF3F5: test eax, eax
  loc_00DFF3F7: fnclex
  loc_00DFF3F9: jge 00DFF410h
  loc_00DFF3FB: mov edx, var_84
  loc_00DFF401: push 00000020h
  loc_00DFF403: push 006DEB78h
  loc_00DFF408: push edx
  loc_00DFF409: push eax
  loc_00DFF40A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF410: lea eax, var_40
  loc_00DFF413: lea ecx, var_30
  loc_00DFF416: push eax
  loc_00DFF417: push ecx
  loc_00DFF418: push 00000002h
  loc_00DFF41A: call [00401038h] ; __vbaFreeVarList
  loc_00DFF420: mov edx, [esi+0000014Ch]
  loc_00DFF426: mov ecx, var_84
  loc_00DFF42C: mov eax, 00000009h
  loc_00DFF431: mov var_28, 00000000h
  loc_00DFF438: mov var_30, eax
  loc_00DFF43B: mov var_48, edx
  loc_00DFF43E: mov var_50, eax
  loc_00DFF441: mov edx, [ecx]
  loc_00DFF443: push ecx
  loc_00DFF444: mov ecx, esp
  loc_00DFF446: sub esp, 00000010h
  loc_00DFF449: mov [ecx], eax
  loc_00DFF44B: mov eax, var_2C
  loc_00DFF44E: mov [ecx+00000004h], eax
  loc_00DFF451: mov eax, var_28
  loc_00DFF454: mov [ecx+00000008h], eax
  loc_00DFF457: mov eax, var_24
  loc_00DFF45A: mov [ecx+0000000Ch], eax
  loc_00DFF45D: mov eax, var_50
  loc_00DFF460: mov ecx, esp
  loc_00DFF462: push 006DEC6Ch ; "PictureNormal"
  loc_00DFF467: mov [ecx], eax
  loc_00DFF469: mov eax, var_4C
  loc_00DFF46C: mov [ecx+00000004h], eax
  loc_00DFF46F: mov eax, var_48
  loc_00DFF472: mov [ecx+00000008h], eax
  loc_00DFF475: mov eax, var_44
  loc_00DFF478: mov [ecx+0000000Ch], eax
  loc_00DFF47B: mov ecx, var_84
  loc_00DFF481: push ecx
  loc_00DFF482: call [edx+00000020h]
  loc_00DFF485: test eax, eax
  loc_00DFF487: fnclex
  loc_00DFF489: jge 00DFF4A0h
  loc_00DFF48B: mov edx, var_84
  loc_00DFF491: push 00000020h
  loc_00DFF493: push 006DEB78h
  loc_00DFF498: push edx
  loc_00DFF499: push eax
  loc_00DFF49A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF4A0: lea ecx, var_30
  loc_00DFF4A3: call [00401024h] ; __vbaFreeVar
  loc_00DFF4A9: mov ecx, [esi+00000150h]
  loc_00DFF4AF: mov edx, var_84
  loc_00DFF4B5: mov eax, 00000009h
  loc_00DFF4BA: mov var_28, 00000000h
  loc_00DFF4C1: mov var_30, eax
  loc_00DFF4C4: mov var_48, ecx
  loc_00DFF4C7: mov var_50, eax
  loc_00DFF4CA: mov ecx, [edx]
  loc_00DFF4CC: sub esp, 00000010h
  loc_00DFF4CF: mov edx, esp
  loc_00DFF4D1: sub esp, 00000010h
  loc_00DFF4D4: mov [edx], eax
  loc_00DFF4D6: mov eax, var_2C
  loc_00DFF4D9: mov [edx+00000004h], eax
  loc_00DFF4DC: mov eax, var_28
  loc_00DFF4DF: mov [edx+00000008h], eax
  loc_00DFF4E2: mov eax, var_24
  loc_00DFF4E5: mov [edx+0000000Ch], eax
  loc_00DFF4E8: mov eax, var_50
  loc_00DFF4EB: mov edx, esp
  loc_00DFF4ED: push 006DEC8Ch ; "PictureHot"
  loc_00DFF4F2: mov [edx], eax
  loc_00DFF4F4: mov eax, var_4C
  loc_00DFF4F7: mov [edx+00000004h], eax
  loc_00DFF4FA: mov eax, var_48
  loc_00DFF4FD: mov [edx+00000008h], eax
  loc_00DFF500: mov eax, var_44
  loc_00DFF503: mov [edx+0000000Ch], eax
  loc_00DFF506: mov edx, var_84
  loc_00DFF50C: push edx
  loc_00DFF50D: call [ecx+00000020h]
  loc_00DFF510: test eax, eax
  loc_00DFF512: fnclex
  loc_00DFF514: jge 00DFF52Bh
  loc_00DFF516: mov ecx, var_84
  loc_00DFF51C: push 00000020h
  loc_00DFF51E: push 006DEB78h
  loc_00DFF523: push ecx
  loc_00DFF524: push eax
  loc_00DFF525: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF52B: lea ecx, var_30
  loc_00DFF52E: call [00401024h] ; __vbaFreeVar
  loc_00DFF534: mov edx, [esi+00000154h]
  loc_00DFF53A: mov ecx, var_84
  loc_00DFF540: mov eax, 00000009h
  loc_00DFF545: mov var_28, 00000000h
  loc_00DFF54C: mov var_30, eax
  loc_00DFF54F: mov var_48, edx
  loc_00DFF552: mov var_50, eax
  loc_00DFF555: mov edx, [ecx]
  loc_00DFF557: sub esp, 00000010h
  loc_00DFF55A: mov ecx, esp
  loc_00DFF55C: sub esp, 00000010h
  loc_00DFF55F: mov [ecx], eax
  loc_00DFF561: mov eax, var_2C
  loc_00DFF564: mov [ecx+00000004h], eax
  loc_00DFF567: mov eax, var_28
  loc_00DFF56A: mov [ecx+00000008h], eax
  loc_00DFF56D: mov eax, var_24
  loc_00DFF570: mov [ecx+0000000Ch], eax
  loc_00DFF573: mov eax, var_50
  loc_00DFF576: mov ecx, esp
  loc_00DFF578: push 006DECA8h ; "PictureDown"
  loc_00DFF57D: mov [ecx], eax
  loc_00DFF57F: mov eax, var_4C
  loc_00DFF582: mov [ecx+00000004h], eax
  loc_00DFF585: mov eax, var_48
  loc_00DFF588: mov [ecx+00000008h], eax
  loc_00DFF58B: mov eax, var_44
  loc_00DFF58E: mov [ecx+0000000Ch], eax
  loc_00DFF591: mov ecx, var_84
  loc_00DFF597: push ecx
  loc_00DFF598: call [edx+00000020h]
  loc_00DFF59B: test eax, eax
  loc_00DFF59D: fnclex
  loc_00DFF59F: jge 00DFF5B6h
  loc_00DFF5A1: mov edx, var_84
  loc_00DFF5A7: push 00000020h
  loc_00DFF5A9: push 006DEB78h
  loc_00DFF5AE: push edx
  loc_00DFF5AF: push eax
  loc_00DFF5B0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF5B6: lea ecx, var_30
  loc_00DFF5B9: call [00401024h] ; __vbaFreeVar
  loc_00DFF5BF: mov ecx, [esi+00000164h]
  loc_00DFF5C5: mov edx, var_84
  loc_00DFF5CB: mov eax, 00000003h
  loc_00DFF5D0: mov var_48, ecx
  loc_00DFF5D3: mov var_50, eax
  loc_00DFF5D6: mov ecx, [edx]
  loc_00DFF5D8: sub esp, 00000010h
  loc_00DFF5DB: mov edx, esp
  loc_00DFF5DD: sub esp, 00000010h
  loc_00DFF5E0: mov [edx], eax
  loc_00DFF5E2: mov eax, 00000001h
  loc_00DFF5E7: mov [edx+00000004h], edi
  loc_00DFF5EA: mov [edx+00000008h], eax
  loc_00DFF5ED: mov eax, esp
  loc_00DFF5EF: push 006DEE54h ; "PictureAlign"
  loc_00DFF5F4: mov [edx+0000000Ch], ebx
  loc_00DFF5F7: mov edx, var_50
  loc_00DFF5FA: mov [eax], edx
  loc_00DFF5FC: mov edx, var_4C
  loc_00DFF5FF: mov [eax+00000004h], edx
  loc_00DFF602: mov edx, var_48
  loc_00DFF605: mov [eax+00000008h], edx
  loc_00DFF608: mov edx, var_44
  loc_00DFF60B: mov [eax+0000000Ch], edx
  loc_00DFF60E: mov eax, var_84
  loc_00DFF614: push eax
  loc_00DFF615: call [ecx+00000020h]
  loc_00DFF618: test eax, eax
  loc_00DFF61A: fnclex
  loc_00DFF61C: jge 00DFF633h
  loc_00DFF61E: mov ecx, var_84
  loc_00DFF624: push 00000020h
  loc_00DFF626: push 006DEB78h
  loc_00DFF62B: push ecx
  loc_00DFF62C: push eax
  loc_00DFF62D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF633: mov edx, [esi+00000168h]
  loc_00DFF639: mov ecx, var_84
  loc_00DFF63F: mov eax, 00000003h
  loc_00DFF644: mov var_48, edx
  loc_00DFF647: mov var_50, eax
  loc_00DFF64A: mov edx, [ecx]
  loc_00DFF64C: sub esp, 00000010h
  loc_00DFF64F: mov ecx, esp
  loc_00DFF651: sub esp, 00000010h
  loc_00DFF654: mov [ecx], eax
  loc_00DFF656: mov eax, 00000001h
  loc_00DFF65B: mov [ecx+00000004h], edi
  loc_00DFF65E: mov [ecx+00000008h], eax
  loc_00DFF661: mov eax, esp
  loc_00DFF663: push 006DED90h ; "PictureEffectOnOver"
  loc_00DFF668: mov [ecx+0000000Ch], ebx
  loc_00DFF66B: mov ecx, var_50
  loc_00DFF66E: mov [eax], ecx
  loc_00DFF670: mov ecx, var_4C
  loc_00DFF673: mov [eax+00000004h], ecx
  loc_00DFF676: mov ecx, var_48
  loc_00DFF679: mov [eax+00000008h], ecx
  loc_00DFF67C: mov ecx, var_44
  loc_00DFF67F: mov [eax+0000000Ch], ecx
  loc_00DFF682: mov eax, var_84
  loc_00DFF688: push eax
  loc_00DFF689: call [edx+00000020h]
  loc_00DFF68C: test eax, eax
  loc_00DFF68E: fnclex
  loc_00DFF690: jge 00DFF6A7h
  loc_00DFF692: mov ecx, var_84
  loc_00DFF698: push 00000020h
  loc_00DFF69A: push 006DEB78h
  loc_00DFF69F: push ecx
  loc_00DFF6A0: push eax
  loc_00DFF6A1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF6A7: mov edx, [esi+0000016Ch]
  loc_00DFF6AD: mov ecx, var_84
  loc_00DFF6B3: mov eax, 00000003h
  loc_00DFF6B8: mov var_48, edx
  loc_00DFF6BB: mov var_50, eax
  loc_00DFF6BE: mov edx, [ecx]
  loc_00DFF6C0: sub esp, 00000010h
  loc_00DFF6C3: mov ecx, esp
  loc_00DFF6C5: sub esp, 00000010h
  loc_00DFF6C8: mov [ecx], eax
  loc_00DFF6CA: mov eax, 00000002h
  loc_00DFF6CF: mov [ecx+00000004h], edi
  loc_00DFF6D2: mov [ecx+00000008h], eax
  loc_00DFF6D5: mov eax, esp
  loc_00DFF6D7: push 006DEDBCh ; "PictureEffectOnDown"
  loc_00DFF6DC: mov [ecx+0000000Ch], ebx
  loc_00DFF6DF: mov ecx, var_50
  loc_00DFF6E2: mov [eax], ecx
  loc_00DFF6E4: mov ecx, var_4C
  loc_00DFF6E7: mov [eax+00000004h], ecx
  loc_00DFF6EA: mov ecx, var_48
  loc_00DFF6ED: mov [eax+00000008h], ecx
  loc_00DFF6F0: mov ecx, var_44
  loc_00DFF6F3: mov [eax+0000000Ch], ecx
  loc_00DFF6F6: mov eax, var_84
  loc_00DFF6FC: push eax
  loc_00DFF6FD: call [edx+00000020h]
  loc_00DFF700: test eax, eax
  loc_00DFF702: fnclex
  loc_00DFF704: jge 00DFF71Bh
  loc_00DFF706: mov ecx, var_84
  loc_00DFF70C: push 00000020h
  loc_00DFF70E: push 006DEB78h
  loc_00DFF713: push ecx
  loc_00DFF714: push eax
  loc_00DFF715: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF71B: mov dx, [esi+00000170h]
  loc_00DFF722: mov ecx, var_84
  loc_00DFF728: mov eax, 0000000Bh
  loc_00DFF72D: mov var_48, dx
  loc_00DFF731: mov var_50, eax
  loc_00DFF734: mov edx, [ecx]
  loc_00DFF736: sub esp, 00000010h
  loc_00DFF739: mov ecx, esp
  loc_00DFF73B: sub esp, 00000010h
  loc_00DFF73E: mov [ecx], eax
  loc_00DFF740: xor eax, eax
  loc_00DFF742: mov [ecx+00000004h], edi
  loc_00DFF745: mov [ecx+00000008h], eax
  loc_00DFF748: mov eax, esp
  loc_00DFF74A: push 006DED64h ; "PicturePushOnHover"
  loc_00DFF74F: mov [ecx+0000000Ch], ebx
  loc_00DFF752: mov ecx, var_50
  loc_00DFF755: mov [eax], ecx
  loc_00DFF757: mov ecx, var_4C
  loc_00DFF75A: mov [eax+00000004h], ecx
  loc_00DFF75D: mov ecx, var_48
  loc_00DFF760: mov [eax+00000008h], ecx
  loc_00DFF763: mov ecx, var_44
  loc_00DFF766: mov [eax+0000000Ch], ecx
  loc_00DFF769: mov eax, var_84
  loc_00DFF76F: push eax
  loc_00DFF770: call [edx+00000020h]
  loc_00DFF773: test eax, eax
  loc_00DFF775: fnclex
  loc_00DFF777: jge 00DFF78Eh
  loc_00DFF779: mov ecx, var_84
  loc_00DFF77F: push 00000020h
  loc_00DFF781: push 006DEB78h
  loc_00DFF786: push ecx
  loc_00DFF787: push eax
  loc_00DFF788: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF78E: lea ecx, var_50
  loc_00DFF791: call [00401024h] ; __vbaFreeVar
  loc_00DFF797: mov dx, [esi+00000160h]
  loc_00DFF79E: mov ecx, var_84
  loc_00DFF7A4: mov eax, 0000000Bh
  loc_00DFF7A9: mov var_48, dx
  loc_00DFF7AD: mov var_50, eax
  loc_00DFF7B0: mov edx, [ecx]
  loc_00DFF7B2: sub esp, 00000010h
  loc_00DFF7B5: mov ecx, esp
  loc_00DFF7B7: sub esp, 00000010h
  loc_00DFF7BA: mov [ecx], eax
  loc_00DFF7BC: xor eax, eax
  loc_00DFF7BE: mov [ecx+00000004h], edi
  loc_00DFF7C1: mov [ecx+00000008h], eax
  loc_00DFF7C4: mov eax, esp
  loc_00DFF7C6: push 006DECC4h ; "PictureShadow"
  loc_00DFF7CB: mov [ecx+0000000Ch], ebx
  loc_00DFF7CE: mov ecx, var_50
  loc_00DFF7D1: mov [eax], ecx
  loc_00DFF7D3: mov ecx, var_4C
  loc_00DFF7D6: mov [eax+00000004h], ecx
  loc_00DFF7D9: mov ecx, var_48
  loc_00DFF7DC: mov [eax+00000008h], ecx
  loc_00DFF7DF: mov ecx, var_44
  loc_00DFF7E2: mov [eax+0000000Ch], ecx
  loc_00DFF7E5: mov eax, var_84
  loc_00DFF7EB: push eax
  loc_00DFF7EC: call [edx+00000020h]
  loc_00DFF7EF: test eax, eax
  loc_00DFF7F1: fnclex
  loc_00DFF7F3: jge 00DFF80Ah
  loc_00DFF7F5: mov ecx, var_84
  loc_00DFF7FB: push 00000020h
  loc_00DFF7FD: push 006DEB78h
  loc_00DFF802: push ecx
  loc_00DFF803: push eax
  loc_00DFF804: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF80A: lea ecx, var_50
  loc_00DFF80D: call [00401024h] ; __vbaFreeVar
  loc_00DFF813: mov dl, [esi+00000158h]
  loc_00DFF819: mov ecx, var_84
  loc_00DFF81F: mov var_48, dl
  loc_00DFF822: mov var_50, 00000011h
  loc_00DFF829: mov edx, [ecx]
  loc_00DFF82B: sub esp, 00000010h
  loc_00DFF82E: mov ecx, esp
  loc_00DFF830: mov eax, 00000002h
  loc_00DFF835: sub esp, 00000010h
  loc_00DFF838: mov [ecx], eax
  loc_00DFF83A: mov eax, 000000FFh
  loc_00DFF83F: mov [ecx+00000004h], edi
  loc_00DFF842: mov [ecx+00000008h], eax
  loc_00DFF845: mov eax, esp
  loc_00DFF847: push 006DECE4h ; "PictureOpacity"
  loc_00DFF84C: mov [ecx+0000000Ch], ebx
  loc_00DFF84F: mov ecx, var_50
  loc_00DFF852: mov [eax], ecx
  loc_00DFF854: mov ecx, var_4C
  loc_00DFF857: mov [eax+00000004h], ecx
  loc_00DFF85A: mov ecx, var_48
  loc_00DFF85D: mov [eax+00000008h], ecx
  loc_00DFF860: mov ecx, var_44
  loc_00DFF863: mov [eax+0000000Ch], ecx
  loc_00DFF866: mov eax, var_84
  loc_00DFF86C: push eax
  loc_00DFF86D: call [edx+00000020h]
  loc_00DFF870: test eax, eax
  loc_00DFF872: fnclex
  loc_00DFF874: jge 00DFF88Bh
  loc_00DFF876: mov ecx, var_84
  loc_00DFF87C: push 00000020h
  loc_00DFF87E: push 006DEB78h
  loc_00DFF883: push ecx
  loc_00DFF884: push eax
  loc_00DFF885: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF88B: mov dl, [esi+00000159h]
  loc_00DFF891: mov ecx, var_84
  loc_00DFF897: mov var_48, dl
  loc_00DFF89A: mov var_50, 00000011h
  loc_00DFF8A1: mov edx, [ecx]
  loc_00DFF8A3: sub esp, 00000010h
  loc_00DFF8A6: mov ecx, esp
  loc_00DFF8A8: mov eax, 00000002h
  loc_00DFF8AD: mov var_58, 000000FFh
  loc_00DFF8B4: sub esp, 00000010h
  loc_00DFF8B7: mov [ecx], eax
  loc_00DFF8B9: mov eax, var_58
  loc_00DFF8BC: mov [ecx+00000004h], edi
  loc_00DFF8BF: mov [ecx+00000008h], eax
  loc_00DFF8C2: mov eax, var_50
  loc_00DFF8C5: mov [ecx+0000000Ch], ebx
  loc_00DFF8C8: mov ecx, esp
  loc_00DFF8CA: push 006DED08h ; "PictureOpacityOnOver"
  loc_00DFF8CF: mov [ecx], eax
  loc_00DFF8D1: mov eax, var_4C
  loc_00DFF8D4: mov [ecx+00000004h], eax
  loc_00DFF8D7: mov eax, var_48
  loc_00DFF8DA: mov [ecx+00000008h], eax
  loc_00DFF8DD: mov eax, var_44
  loc_00DFF8E0: mov [ecx+0000000Ch], eax
  loc_00DFF8E3: mov ecx, var_84
  loc_00DFF8E9: push ecx
  loc_00DFF8EA: call [edx+00000020h]
  loc_00DFF8ED: test eax, eax
  loc_00DFF8EF: fnclex
  loc_00DFF8F1: jge 00DFF908h
  loc_00DFF8F3: mov edx, var_84
  loc_00DFF8F9: push 00000020h
  loc_00DFF8FB: push 006DEB78h
  loc_00DFF900: push edx
  loc_00DFF901: push eax
  loc_00DFF902: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF908: mov ecx, [esi+0000015Ch]
  loc_00DFF90E: mov edx, var_84
  loc_00DFF914: mov eax, 00000003h
  loc_00DFF919: mov var_48, ecx
  loc_00DFF91C: mov var_50, eax
  loc_00DFF91F: mov ecx, [edx]
  loc_00DFF921: sub esp, 00000010h
  loc_00DFF924: mov edx, esp
  loc_00DFF926: sub esp, 00000010h
  loc_00DFF929: mov [edx], eax
  loc_00DFF92B: xor eax, eax
  loc_00DFF92D: mov [edx+00000004h], edi
  loc_00DFF930: mov [edx+00000008h], eax
  loc_00DFF933: mov eax, esp
  loc_00DFF935: push 006DED38h ; "DisabledPictureMode"
  loc_00DFF93A: mov [edx+0000000Ch], ebx
  loc_00DFF93D: mov edx, var_50
  loc_00DFF940: mov [eax], edx
  loc_00DFF942: mov edx, var_4C
  loc_00DFF945: mov [eax+00000004h], edx
  loc_00DFF948: mov edx, var_48
  loc_00DFF94B: mov [eax+00000008h], edx
  loc_00DFF94E: mov edx, var_44
  loc_00DFF951: mov [eax+0000000Ch], edx
  loc_00DFF954: mov eax, var_84
  loc_00DFF95A: push eax
  loc_00DFF95B: call [ecx+00000020h]
  loc_00DFF95E: test eax, eax
  loc_00DFF960: fnclex
  loc_00DFF962: jge 00DFF979h
  loc_00DFF964: mov ecx, var_84
  loc_00DFF96A: push 00000020h
  loc_00DFF96C: push 006DEB78h
  loc_00DFF971: push ecx
  loc_00DFF972: push eax
  loc_00DFF973: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF979: mov edx, [esi+00000068h]
  loc_00DFF97C: mov ecx, var_84
  loc_00DFF982: mov var_48, edx
  loc_00DFF985: mov var_50, 00000003h
  loc_00DFF98C: mov edx, [ecx]
  loc_00DFF98E: sub esp, 00000010h
  loc_00DFF991: mov ecx, esp
  loc_00DFF993: mov eax, 00000008h
  loc_00DFF998: sub esp, 00000010h
  loc_00DFF99B: mov [ecx], eax
  loc_00DFF99D: xor eax, eax
  loc_00DFF99F: mov [ecx+00000004h], edi
  loc_00DFF9A2: mov [ecx+00000008h], eax
  loc_00DFF9A5: mov eax, esp
  loc_00DFF9A7: push 006DEE20h ; "CaptionEffects"
  loc_00DFF9AC: mov [ecx+0000000Ch], ebx
  loc_00DFF9AF: mov ecx, var_50
  loc_00DFF9B2: mov [eax], ecx
  loc_00DFF9B4: mov ecx, var_4C
  loc_00DFF9B7: mov [eax+00000004h], ecx
  loc_00DFF9BA: mov ecx, var_48
  loc_00DFF9BD: mov [eax+00000008h], ecx
  loc_00DFF9C0: mov ecx, var_44
  loc_00DFF9C3: mov [eax+0000000Ch], ecx
  loc_00DFF9C6: mov eax, var_84
  loc_00DFF9CC: push eax
  loc_00DFF9CD: call [edx+00000020h]
  loc_00DFF9D0: test eax, eax
  loc_00DFF9D2: fnclex
  loc_00DFF9D4: jge 00DFF9EBh
  loc_00DFF9D6: mov ecx, var_84
  loc_00DFF9DC: push 00000020h
  loc_00DFF9DE: push 006DEB78h
  loc_00DFF9E3: push ecx
  loc_00DFF9E4: push eax
  loc_00DFF9E5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFF9EB: mov dx, [esi+0000009Ch]
  loc_00DFF9F2: mov ecx, var_84
  loc_00DFF9F8: mov eax, 0000000Bh
  loc_00DFF9FD: mov var_48, dx
  loc_00DFFA01: mov var_50, eax
  loc_00DFFA04: mov edx, [ecx]
  loc_00DFFA06: sub esp, 00000010h
  loc_00DFFA09: mov var_58, FFFFFFFFh
  loc_00DFFA10: mov ecx, esp
  loc_00DFFA12: sub esp, 00000010h
  loc_00DFFA15: mov [ecx], eax
  loc_00DFFA17: mov eax, var_58
  loc_00DFFA1A: mov [ecx+00000004h], edi
  loc_00DFFA1D: mov [ecx+00000008h], eax
  loc_00DFFA20: mov eax, var_50
  loc_00DFFA23: mov [ecx+0000000Ch], ebx
  loc_00DFFA26: mov ecx, esp
  loc_00DFFA28: push 006DEE00h ; "UseMaskColor"
  loc_00DFFA2D: mov [ecx], eax
  loc_00DFFA2F: mov eax, var_4C
  loc_00DFFA32: mov [ecx+00000004h], eax
  loc_00DFFA35: mov eax, var_48
  loc_00DFFA38: mov [ecx+00000008h], eax
  loc_00DFFA3B: mov eax, var_44
  loc_00DFFA3E: mov [ecx+0000000Ch], eax
  loc_00DFFA41: mov ecx, var_84
  loc_00DFFA47: push ecx
  loc_00DFFA48: call [edx+00000020h]
  loc_00DFFA4B: test eax, eax
  loc_00DFFA4D: fnclex
  loc_00DFFA4F: jge 00DFFA66h
  loc_00DFFA51: mov edx, var_84
  loc_00DFFA57: push 00000020h
  loc_00DFFA59: push 006DEB78h
  loc_00DFFA5E: push edx
  loc_00DFFA5F: push eax
  loc_00DFFA60: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFA66: lea ecx, var_50
  loc_00DFFA69: call [00401024h] ; __vbaFreeVar
  loc_00DFFA6F: mov ecx, [esi+000000A0h]
  loc_00DFFA75: mov edx, var_84
  loc_00DFFA7B: mov eax, 00000003h
  loc_00DFFA80: mov var_48, ecx
  loc_00DFFA83: mov var_50, eax
  loc_00DFFA86: mov ecx, [edx]
  loc_00DFFA88: sub esp, 00000010h
  loc_00DFFA8B: mov edx, esp
  loc_00DFFA8D: sub esp, 00000010h
  loc_00DFFA90: mov [edx], eax
  loc_00DFFA92: mov eax, 00E0E0E0h
  loc_00DFFA97: mov [edx+00000004h], edi
  loc_00DFFA9A: mov [edx+00000008h], eax
  loc_00DFFA9D: mov eax, esp
  loc_00DFFA9F: push 006DEDE8h ; "MaskColor"
  loc_00DFFAA4: mov [edx+0000000Ch], ebx
  loc_00DFFAA7: mov edx, var_50
  loc_00DFFAAA: mov [eax], edx
  loc_00DFFAAC: mov edx, var_4C
  loc_00DFFAAF: mov [eax+00000004h], edx
  loc_00DFFAB2: mov edx, var_48
  loc_00DFFAB5: mov [eax+00000008h], edx
  loc_00DFFAB8: mov edx, var_44
  loc_00DFFABB: mov [eax+0000000Ch], edx
  loc_00DFFABE: mov eax, var_84
  loc_00DFFAC4: push eax
  loc_00DFFAC5: call [ecx+00000020h]
  loc_00DFFAC8: test eax, eax
  loc_00DFFACA: fnclex
  loc_00DFFACC: jge 00DFFAE3h
  loc_00DFFACE: mov ecx, var_84
  loc_00DFFAD4: push 00000020h
  loc_00DFFAD6: push 006DEB78h
  loc_00DFFADB: push ecx
  loc_00DFFADC: push eax
  loc_00DFFADD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFAE3: mov edx, [esi+00000084h]
  loc_00DFFAE9: mov ecx, var_84
  loc_00DFFAEF: mov eax, 00000003h
  loc_00DFFAF4: mov var_48, edx
  loc_00DFFAF7: mov var_50, eax
  loc_00DFFAFA: mov edx, [ecx]
  loc_00DFFAFC: sub esp, 00000010h
  loc_00DFFAFF: mov var_58, 00000001h
  loc_00DFFB06: mov ecx, esp
  loc_00DFFB08: sub esp, 00000010h
  loc_00DFFB0B: mov [ecx], eax
  loc_00DFFB0D: mov eax, var_58
  loc_00DFFB10: mov [ecx+00000004h], edi
  loc_00DFFB13: mov [ecx+00000008h], eax
  loc_00DFFB16: mov eax, var_50
  loc_00DFFB19: mov [ecx+0000000Ch], ebx
  loc_00DFFB1C: mov ecx, esp
  loc_00DFFB1E: push 006DEE74h ; "CaptionAlign"
  loc_00DFFB23: mov [ecx], eax
  loc_00DFFB25: mov eax, var_4C
  loc_00DFFB28: mov [ecx+00000004h], eax
  loc_00DFFB2B: mov eax, var_48
  loc_00DFFB2E: mov [ecx+00000008h], eax
  loc_00DFFB31: mov eax, var_44
  loc_00DFFB34: mov [ecx+0000000Ch], eax
  loc_00DFFB37: mov ecx, var_84
  loc_00DFFB3D: push ecx
  loc_00DFFB3E: call [edx+00000020h]
  loc_00DFFB41: test eax, eax
  loc_00DFFB43: fnclex
  loc_00DFFB45: jge 00DFFB5Ch
  loc_00DFFB47: mov edx, var_84
  loc_00DFFB4D: push 00000020h
  loc_00DFFB4F: push 006DEB78h
  loc_00DFFB54: push edx
  loc_00DFFB55: push eax
  loc_00DFFB56: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFB5C: mov ecx, [esi+000000F8h]
  loc_00DFFB62: mov edx, var_84
  loc_00DFFB68: mov eax, 00000008h
  loc_00DFFB6D: mov var_48, ecx
  loc_00DFFB70: mov var_50, eax
  loc_00DFFB73: mov ecx, [edx]
  loc_00DFFB75: sub esp, 00000010h
  loc_00DFFB78: mov edx, esp
  loc_00DFFB7A: sub esp, 00000010h
  loc_00DFFB7D: mov [edx], eax
  loc_00DFFB7F: xor eax, eax
  loc_00DFFB81: mov [edx+00000004h], edi
  loc_00DFFB84: mov [edx+00000008h], eax
  loc_00DFFB87: mov eax, esp
  loc_00DFFB89: push 006DEF18h ; "ToolTip"
  loc_00DFFB8E: mov [edx+0000000Ch], ebx
  loc_00DFFB91: mov edx, var_50
  loc_00DFFB94: mov [eax], edx
  loc_00DFFB96: mov edx, var_4C
  loc_00DFFB99: mov [eax+00000004h], edx
  loc_00DFFB9C: mov edx, var_48
  loc_00DFFB9F: mov [eax+00000008h], edx
  loc_00DFFBA2: mov edx, var_44
  loc_00DFFBA5: mov [eax+0000000Ch], edx
  loc_00DFFBA8: mov eax, var_84
  loc_00DFFBAE: push eax
  loc_00DFFBAF: call [ecx+00000020h]
  loc_00DFFBB2: test eax, eax
  loc_00DFFBB4: fnclex
  loc_00DFFBB6: jge 00DFFBCDh
  loc_00DFFBB8: mov ecx, var_84
  loc_00DFFBBE: push 00000020h
  loc_00DFFBC0: push 006DEB78h
  loc_00DFFBC5: push ecx
  loc_00DFFBC6: push eax
  loc_00DFFBC7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFBCD: mov edx, [esi+00000104h]
  loc_00DFFBD3: mov ecx, var_84
  loc_00DFFBD9: mov eax, 00000003h
  loc_00DFFBDE: mov var_48, edx
  loc_00DFFBE1: mov var_50, eax
  loc_00DFFBE4: mov edx, [ecx]
  loc_00DFFBE6: sub esp, 00000010h
  loc_00DFFBE9: mov ecx, esp
  loc_00DFFBEB: sub esp, 00000010h
  loc_00DFFBEE: mov [ecx], eax
  loc_00DFFBF0: xor eax, eax
  loc_00DFFBF2: mov [ecx+00000004h], edi
  loc_00DFFBF5: mov [ecx+00000008h], eax
  loc_00DFFBF8: mov eax, esp
  loc_00DFFBFA: push 006DEF48h ; "TooltipType"
  loc_00DFFBFF: mov [ecx+0000000Ch], ebx
  loc_00DFFC02: mov ecx, var_50
  loc_00DFFC05: mov [eax], ecx
  loc_00DFFC07: mov ecx, var_4C
  loc_00DFFC0A: mov [eax+00000004h], ecx
  loc_00DFFC0D: mov ecx, var_48
  loc_00DFFC10: mov [eax+00000008h], ecx
  loc_00DFFC13: mov ecx, var_44
  loc_00DFFC16: mov [eax+0000000Ch], ecx
  loc_00DFFC19: mov eax, var_84
  loc_00DFFC1F: push eax
  loc_00DFFC20: call [edx+00000020h]
  loc_00DFFC23: test eax, eax
  loc_00DFFC25: fnclex
  loc_00DFFC27: jge 00DFFC3Eh
  loc_00DFFC29: mov ecx, var_84
  loc_00DFFC2F: push 00000020h
  loc_00DFFC31: push 006DEB78h
  loc_00DFFC36: push ecx
  loc_00DFFC37: push eax
  loc_00DFFC38: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFC3E: mov edx, [esi+00000100h]
  loc_00DFFC44: mov ecx, var_84
  loc_00DFFC4A: mov eax, 00000003h
  loc_00DFFC4F: mov var_48, edx
  loc_00DFFC52: mov var_50, eax
  loc_00DFFC55: mov edx, [ecx]
  loc_00DFFC57: sub esp, 00000010h
  loc_00DFFC5A: mov ecx, esp
  loc_00DFFC5C: sub esp, 00000010h
  loc_00DFFC5F: mov [ecx], eax
  loc_00DFFC61: xor eax, eax
  loc_00DFFC63: mov [ecx+00000004h], edi
  loc_00DFFC66: mov [ecx+00000008h], eax
  loc_00DFFC69: mov eax, esp
  loc_00DFFC6B: push 006DEF2Ch ; "TooltipIcon"
  loc_00DFFC70: mov [ecx+0000000Ch], ebx
  loc_00DFFC73: mov ecx, var_50
  loc_00DFFC76: mov [eax], ecx
  loc_00DFFC78: mov ecx, var_4C
  loc_00DFFC7B: mov [eax+00000004h], ecx
  loc_00DFFC7E: mov ecx, var_48
  loc_00DFFC81: mov [eax+00000008h], ecx
  loc_00DFFC84: mov ecx, var_44
  loc_00DFFC87: mov [eax+0000000Ch], ecx
  loc_00DFFC8A: mov eax, var_84
  loc_00DFFC90: push eax
  loc_00DFFC91: call [edx+00000020h]
  loc_00DFFC94: test eax, eax
  loc_00DFFC96: fnclex
  loc_00DFFC98: jge 00DFFCAFh
  loc_00DFFC9A: mov ecx, var_84
  loc_00DFFCA0: push 00000020h
  loc_00DFFCA2: push 006DEB78h
  loc_00DFFCA7: push ecx
  loc_00DFFCA8: push eax
  loc_00DFFCA9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFCAF: mov edx, [esi+000000FCh]
  loc_00DFFCB5: mov ecx, var_84
  loc_00DFFCBB: mov eax, 00000008h
  loc_00DFFCC0: mov var_48, edx
  loc_00DFFCC3: mov var_50, eax
  loc_00DFFCC6: mov edx, [ecx]
  loc_00DFFCC8: sub esp, 00000010h
  loc_00DFFCCB: mov ecx, esp
  loc_00DFFCCD: sub esp, 00000010h
  loc_00DFFCD0: mov [ecx], eax
  loc_00DFFCD2: xor eax, eax
  loc_00DFFCD4: mov [ecx+00000004h], edi
  loc_00DFFCD7: mov [ecx+00000008h], eax
  loc_00DFFCDA: mov eax, esp
  loc_00DFFCDC: push 006DEEF8h ; "TooltipTitle"
  loc_00DFFCE1: mov [ecx+0000000Ch], ebx
  loc_00DFFCE4: mov ecx, var_50
  loc_00DFFCE7: mov [eax], ecx
  loc_00DFFCE9: mov ecx, var_4C
  loc_00DFFCEC: mov [eax+00000004h], ecx
  loc_00DFFCEF: mov ecx, var_48
  loc_00DFFCF2: mov [eax+00000008h], ecx
  loc_00DFFCF5: mov ecx, var_44
  loc_00DFFCF8: mov [eax+0000000Ch], ecx
  loc_00DFFCFB: mov eax, var_84
  loc_00DFFD01: push eax
  loc_00DFFD02: call [edx+00000020h]
  loc_00DFFD05: test eax, eax
  loc_00DFFD07: fnclex
  loc_00DFFD09: jge 00DFFD20h
  loc_00DFFD0B: mov ecx, var_84
  loc_00DFFD11: push 00000020h
  loc_00DFFD13: push 006DEB78h
  loc_00DFFD18: push ecx
  loc_00DFFD19: push eax
  loc_00DFFD1A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFD20: mov edx, [esi]
  loc_00DFFD22: lea eax, var_6C
  loc_00DFFD25: lea ecx, var_68
  loc_00DFFD28: push eax
  loc_00DFFD29: push ecx
  loc_00DFFD2A: push 80000018h
  loc_00DFFD2F: push esi
  loc_00DFFD30: mov var_68, 00000000h
  loc_00DFFD37: call [edx+0000090Ch]
  loc_00DFFD3D: mov edx, [esi+00000108h]
  loc_00DFFD43: mov ecx, var_84
  loc_00DFFD49: mov eax, 00000003h
  loc_00DFFD4E: mov var_48, edx
  loc_00DFFD51: mov var_50, eax
  loc_00DFFD54: mov edx, [ecx]
  loc_00DFFD56: sub esp, 00000010h
  loc_00DFFD59: mov ecx, esp
  loc_00DFFD5B: sub esp, 00000010h
  loc_00DFFD5E: mov [ecx], eax
  loc_00DFFD60: mov eax, var_6C
  loc_00DFFD63: mov [ecx+00000004h], edi
  loc_00DFFD66: mov [ecx+00000008h], eax
  loc_00DFFD69: mov eax, esp
  loc_00DFFD6B: push 006DEF64h ; "TooltipBackColor"
  loc_00DFFD70: mov [ecx+0000000Ch], ebx
  loc_00DFFD73: mov ecx, var_50
  loc_00DFFD76: mov [eax], ecx
  loc_00DFFD78: mov ecx, var_4C
  loc_00DFFD7B: mov [eax+00000004h], ecx
  loc_00DFFD7E: mov ecx, var_48
  loc_00DFFD81: mov [eax+00000008h], ecx
  loc_00DFFD84: mov ecx, var_44
  loc_00DFFD87: mov [eax+0000000Ch], ecx
  loc_00DFFD8A: mov eax, var_84
  loc_00DFFD90: push eax
  loc_00DFFD91: call [edx+00000020h]
  loc_00DFFD94: test eax, eax
  loc_00DFFD96: fnclex
  loc_00DFFD98: jge 00DFFDAFh
  loc_00DFFD9A: mov ecx, var_84
  loc_00DFFDA0: push 00000020h
  loc_00DFFDA2: push 006DEB78h
  loc_00DFFDA7: push ecx
  loc_00DFFDA8: push eax
  loc_00DFFDA9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFDAF: mov dx, [esi+00000138h]
  loc_00DFFDB6: mov ecx, var_84
  loc_00DFFDBC: mov eax, 0000000Bh
  loc_00DFFDC1: mov var_48, dx
  loc_00DFFDC5: mov var_50, eax
  loc_00DFFDC8: mov edx, [ecx]
  loc_00DFFDCA: sub esp, 00000010h
  loc_00DFFDCD: mov ecx, esp
  loc_00DFFDCF: sub esp, 00000010h
  loc_00DFFDD2: mov [ecx], eax
  loc_00DFFDD4: xor eax, eax
  loc_00DFFDD6: mov [ecx+00000004h], edi
  loc_00DFFDD9: mov [ecx+00000008h], eax
  loc_00DFFDDC: mov eax, esp
  loc_00DFFDDE: push 006DEF8Ch ; "RightToLeft"
  loc_00DFFDE3: mov [ecx+0000000Ch], ebx
  loc_00DFFDE6: mov ecx, var_50
  loc_00DFFDE9: mov [eax], ecx
  loc_00DFFDEB: mov ecx, var_4C
  loc_00DFFDEE: mov [eax+00000004h], ecx
  loc_00DFFDF1: mov ecx, var_48
  loc_00DFFDF4: mov [eax+00000008h], ecx
  loc_00DFFDF7: mov ecx, var_44
  loc_00DFFDFA: mov [eax+0000000Ch], ecx
  loc_00DFFDFD: mov eax, var_84
  loc_00DFFE03: push eax
  loc_00DFFE04: call [edx+00000020h]
  loc_00DFFE07: test eax, eax
  loc_00DFFE09: fnclex
  loc_00DFFE0B: jge 00DFFE22h
  loc_00DFFE0D: mov ecx, var_84
  loc_00DFFE13: push 00000020h
  loc_00DFFE15: push 006DEB78h
  loc_00DFFE1A: push ecx
  loc_00DFFE1B: push eax
  loc_00DFFE1C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFE22: lea ecx, var_50
  loc_00DFFE25: call [00401024h] ; __vbaFreeVar
  loc_00DFFE2B: mov edx, [esi+0000005Ch]
  loc_00DFFE2E: mov ecx, var_84
  loc_00DFFE34: mov eax, 00000003h
  loc_00DFFE39: mov var_48, edx
  loc_00DFFE3C: mov var_50, eax
  loc_00DFFE3F: mov edx, [ecx]
  loc_00DFFE41: sub esp, 00000010h
  loc_00DFFE44: mov ecx, esp
  loc_00DFFE46: sub esp, 00000010h
  loc_00DFFE49: mov [ecx], eax
  loc_00DFFE4B: xor eax, eax
  loc_00DFFE4D: mov [ecx+00000004h], edi
  loc_00DFFE50: mov [ecx+00000008h], eax
  loc_00DFFE53: mov eax, esp
  loc_00DFFE55: push 006DEFA8h ; "DropDownSymbol"
  loc_00DFFE5A: mov [ecx+0000000Ch], ebx
  loc_00DFFE5D: mov ecx, var_50
  loc_00DFFE60: mov [eax], ecx
  loc_00DFFE62: mov ecx, var_4C
  loc_00DFFE65: mov [eax+00000004h], ecx
  loc_00DFFE68: mov ecx, var_48
  loc_00DFFE6B: mov [eax+00000008h], ecx
  loc_00DFFE6E: mov ecx, var_44
  loc_00DFFE71: mov [eax+0000000Ch], ecx
  loc_00DFFE74: mov eax, var_84
  loc_00DFFE7A: push eax
  loc_00DFFE7B: call [edx+00000020h]
  loc_00DFFE7E: test eax, eax
  loc_00DFFE80: fnclex
  loc_00DFFE82: jge 00DFFE99h
  loc_00DFFE84: mov ecx, var_84
  loc_00DFFE8A: push 00000020h
  loc_00DFFE8C: push 006DEB78h
  loc_00DFFE91: push ecx
  loc_00DFFE92: push eax
  loc_00DFFE93: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFE99: mov dx, [esi+00000060h]
  loc_00DFFE9D: mov ecx, var_84
  loc_00DFFEA3: mov eax, 0000000Bh
  loc_00DFFEA8: mov var_48, dx
  loc_00DFFEAC: mov var_50, eax
  loc_00DFFEAF: mov edx, [ecx]
  loc_00DFFEB1: sub esp, 00000010h
  loc_00DFFEB4: mov ecx, esp
  loc_00DFFEB6: sub esp, 00000010h
  loc_00DFFEB9: mov [ecx], eax
  loc_00DFFEBB: xor eax, eax
  loc_00DFFEBD: mov [ecx+00000004h], edi
  loc_00DFFEC0: mov [ecx+00000008h], eax
  loc_00DFFEC3: mov eax, esp
  loc_00DFFEC5: push 006DEED0h ; "DropDownSeparator"
  loc_00DFFECA: mov [ecx+0000000Ch], ebx
  loc_00DFFECD: mov ecx, var_50
  loc_00DFFED0: mov [eax], ecx
  loc_00DFFED2: mov ecx, var_4C
  loc_00DFFED5: mov [eax+00000004h], ecx
  loc_00DFFED8: mov ecx, var_48
  loc_00DFFEDB: mov [eax+00000008h], ecx
  loc_00DFFEDE: mov ecx, var_44
  loc_00DFFEE1: mov [eax+0000000Ch], ecx
  loc_00DFFEE4: mov eax, var_84
  loc_00DFFEEA: push eax
  loc_00DFFEEB: call [edx+00000020h]
  loc_00DFFEEE: test eax, eax
  loc_00DFFEF0: fnclex
  loc_00DFFEF2: jge 00DFFF09h
  loc_00DFFEF4: mov ecx, var_84
  loc_00DFFEFA: push 00000020h
  loc_00DFFEFC: push 006DEB78h
  loc_00DFFF01: push ecx
  loc_00DFFF02: push eax
  loc_00DFFF03: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFF09: lea ecx, var_50
  loc_00DFFF0C: call [00401024h] ; __vbaFreeVar
  loc_00DFFF12: mov edx, [esi+000000CCh]
  loc_00DFFF18: sub esp, 00000010h
  loc_00DFFF1B: mov esi, esp
  loc_00DFFF1D: mov ecx, 00000003h
  loc_00DFFF22: xor eax, eax
  loc_00DFFF24: sub esp, 00000010h
  loc_00DFFF27: mov [esi], ecx
  loc_00DFFF29: mov var_50, ecx
  loc_00DFFF2C: mov var_48, edx
  loc_00DFFF2F: mov edx, var_84
  loc_00DFFF35: mov [esi+00000004h], edi
  loc_00DFFF38: mov edx, [edx]
  loc_00DFFF3A: mov [esi+00000008h], eax
  loc_00DFFF3D: mov eax, esp
  loc_00DFFF3F: push 006DEFE0h ; "ColorScheme"
  loc_00DFFF44: mov [esi+0000000Ch], ebx
  loc_00DFFF47: mov [eax], ecx
  loc_00DFFF49: mov ecx, var_4C
  loc_00DFFF4C: mov [eax+00000004h], ecx
  loc_00DFFF4F: mov ecx, var_48
  loc_00DFFF52: mov [eax+00000008h], ecx
  loc_00DFFF55: mov ecx, var_44
  loc_00DFFF58: mov [eax+0000000Ch], ecx
  loc_00DFFF5B: mov eax, var_84
  loc_00DFFF61: push eax
  loc_00DFFF62: call [edx+00000020h]
  loc_00DFFF65: test eax, eax
  loc_00DFFF67: fnclex
  loc_00DFFF69: jge 00DFFF80h
  loc_00DFFF6B: mov ecx, var_84
  loc_00DFFF71: push 00000020h
  loc_00DFFF73: push 006DEB78h
  loc_00DFFF78: push ecx
  loc_00DFFF79: push eax
  loc_00DFFF7A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFFF80: lea edx, var_84
  loc_00DFFF86: push 00000000h
  loc_00DFFF88: push edx
  loc_00DFFF89: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFFF8F: mov var_4, 00000000h
  loc_00DFFF96: push 00DFFFD2h
  loc_00DFFF9B: jmp 00DFFFC5h
  loc_00DFFF9D: lea eax, var_20
  loc_00DFFFA0: lea ecx, var_1C
  loc_00DFFFA3: push eax
  loc_00DFFFA4: lea edx, var_18
  loc_00DFFFA7: push ecx
  loc_00DFFFA8: push edx
  loc_00DFFFA9: push 00000003h
  loc_00DFFFAB: call [00401048h] ; __vbaFreeObjList
  loc_00DFFFB1: lea eax, var_40
  loc_00DFFFB4: lea ecx, var_30
  loc_00DFFFB7: push eax
  loc_00DFFFB8: push ecx
  loc_00DFFFB9: push 00000002h
  loc_00DFFFBB: call [00401038h] ; __vbaFreeVarList
  loc_00DFFFC1: add esp, 0000001Ch
  loc_00DFFFC4: ret
  loc_00DFFFC5: lea ecx, var_84
  loc_00DFFFCB: call [00401254h] ; __vbaFreeObj
  loc_00DFFFD1: ret
  loc_00DFFFD2: mov eax, Me
  loc_00DFFFD5: push eax
  loc_00DFFFD6: mov edx, [eax]
  loc_00DFFFD8: call [edx+00000008h]
  loc_00DFFFDB: mov eax, var_4
  loc_00DFFFDE: mov ecx, var_14
  loc_00DFFFE1: pop edi
  loc_00DFFFE2: pop esi
  loc_00DFFFE3: mov fs:[00000000h], ecx
  loc_00DFFFEA: pop ebx
  loc_00DFFFEB: mov esp, ebp
  loc_00DFFFED: pop ebp
  loc_00DFFFEE: retn 0008h
End Sub

Private Sub UserControl_UnknownEvent_20 'DFC5C0
  loc_00DFC5C0: push ebp
  loc_00DFC5C1: mov ebp, esp
  loc_00DFC5C3: sub esp, 00000018h
  loc_00DFC5C6: push 00402836h ; __vbaExceptHandler
  loc_00DFC5CB: mov eax, fs:[00000000h]
  loc_00DFC5D1: push eax
  loc_00DFC5D2: mov fs:[00000000h], esp
  loc_00DFC5D9: mov eax, 0000016Ch
  loc_00DFC5DE: call 00402830h ; __vbaChkstk
  loc_00DFC5E3: push ebx
  loc_00DFC5E4: push esi
  loc_00DFC5E5: push edi
  loc_00DFC5E6: mov var_18, esp
  loc_00DFC5E9: mov var_14, 00401950h
  loc_00DFC5F0: mov eax, Me
  loc_00DFC5F3: and eax, 00000001h
  loc_00DFC5F6: mov var_10, eax
  loc_00DFC5F9: mov ecx, Me
  loc_00DFC5FC: and ecx, FFFFFFFEh
  loc_00DFC5FF: mov Me, ecx
  loc_00DFC602: mov var_C, 00000000h
  loc_00DFC609: mov edx, Me
  loc_00DFC60C: mov eax, [edx]
  loc_00DFC60E: mov ecx, Me
  loc_00DFC611: push ecx
  loc_00DFC612: call [eax+00000004h]
  loc_00DFC615: mov var_4, 00000001h
  loc_00DFC61C: mov var_4, 00000002h
  loc_00DFC623: mov edx, Cancel
  loc_00DFC626: mov eax, [edx]
  loc_00DFC628: push eax
  loc_00DFC629: lea ecx, var_84
  loc_00DFC62F: push ecx
  loc_00DFC630: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFC636: mov var_4, 00000003h
  loc_00DFC63D: mov var_58, 00000001h
  loc_00DFC644: mov var_60, 00000003h
  loc_00DFC64B: lea edx, var_40
  loc_00DFC64E: push edx
  loc_00DFC64F: mov eax, 00000010h
  loc_00DFC654: call 00402830h ; __vbaChkstk
  loc_00DFC659: mov eax, esp
  loc_00DFC65B: mov ecx, var_60
  loc_00DFC65E: mov [eax], ecx
  loc_00DFC660: mov edx, var_5C
  loc_00DFC663: mov [eax+00000004h], edx
  loc_00DFC666: mov ecx, var_58
  loc_00DFC669: mov [eax+00000008h], ecx
  loc_00DFC66C: mov edx, var_54
  loc_00DFC66F: mov [eax+0000000Ch], edx
  loc_00DFC672: push 006DEB60h ; "ButtonStyle"
  loc_00DFC677: mov eax, var_84
  loc_00DFC67D: mov ecx, [eax]
  loc_00DFC67F: mov edx, var_84
  loc_00DFC685: push edx
  loc_00DFC686: call [ecx+0000001Ch]
  loc_00DFC689: fnclex
  loc_00DFC68B: mov var_70, eax
  loc_00DFC68E: cmp var_70, 00000000h
  loc_00DFC692: jge 00DFC6B4h
  loc_00DFC694: push 0000001Ch
  loc_00DFC696: push 006DEB78h
  loc_00DFC69B: mov eax, var_84
  loc_00DFC6A1: push eax
  loc_00DFC6A2: mov ecx, var_70
  loc_00DFC6A5: push ecx
  loc_00DFC6A6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC6AC: mov var_A8, eax
  loc_00DFC6B2: jmp 00DFC6BEh
  loc_00DFC6B4: mov var_A8, 00000000h
  loc_00DFC6BE: lea edx, var_40
  loc_00DFC6C1: push edx
  loc_00DFC6C2: call [004011D0h] ; __vbaI4Var
  loc_00DFC6C8: mov ecx, Me
  loc_00DFC6CB: mov [ecx+00000044h], eax
  loc_00DFC6CE: lea ecx, var_40
  loc_00DFC6D1: call [00401024h] ; __vbaFreeVar
  loc_00DFC6D7: mov var_4, 00000004h
  loc_00DFC6DE: mov var_58, 00000000h
  loc_00DFC6E5: mov var_60, 0000000Bh
  loc_00DFC6EC: lea edx, var_40
  loc_00DFC6EF: push edx
  loc_00DFC6F0: mov eax, 00000010h
  loc_00DFC6F5: call 00402830h ; __vbaChkstk
  loc_00DFC6FA: mov eax, esp
  loc_00DFC6FC: mov ecx, var_60
  loc_00DFC6FF: mov [eax], ecx
  loc_00DFC701: mov edx, var_5C
  loc_00DFC704: mov [eax+00000004h], edx
  loc_00DFC707: mov ecx, var_58
  loc_00DFC70A: mov [eax+00000008h], ecx
  loc_00DFC70D: mov edx, var_54
  loc_00DFC710: mov [eax+0000000Ch], edx
  loc_00DFC713: push 006DEB8Ch ; "ShowFocusRect"
  loc_00DFC718: mov eax, var_84
  loc_00DFC71E: mov ecx, [eax]
  loc_00DFC720: mov edx, var_84
  loc_00DFC726: push edx
  loc_00DFC727: call [ecx+0000001Ch]
  loc_00DFC72A: fnclex
  loc_00DFC72C: mov var_70, eax
  loc_00DFC72F: cmp var_70, 00000000h
  loc_00DFC733: jge 00DFC755h
  loc_00DFC735: push 0000001Ch
  loc_00DFC737: push 006DEB78h
  loc_00DFC73C: mov eax, var_84
  loc_00DFC742: push eax
  loc_00DFC743: mov ecx, var_70
  loc_00DFC746: push ecx
  loc_00DFC747: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC74D: mov var_AC, eax
  loc_00DFC753: jmp 00DFC75Fh
  loc_00DFC755: mov var_AC, 00000000h
  loc_00DFC75F: lea edx, var_40
  loc_00DFC762: push edx
  loc_00DFC763: call [004010CCh] ; __vbaBoolVar
  loc_00DFC769: mov ecx, Me
  loc_00DFC76C: mov [ecx+0000006Eh], ax
  loc_00DFC770: lea ecx, var_40
  loc_00DFC773: call [00401024h] ; __vbaFreeVar
  loc_00DFC779: mov var_4, 00000005h
  loc_00DFC780: lea edx, var_28
  loc_00DFC783: push edx
  loc_00DFC784: mov eax, Me
  loc_00DFC787: mov ecx, [eax]
  loc_00DFC789: mov edx, Me
  loc_00DFC78C: push edx
  loc_00DFC78D: call [ecx+000002B0h]
  loc_00DFC793: fnclex
  loc_00DFC795: mov var_70, eax
  loc_00DFC798: cmp var_70, 00000000h
  loc_00DFC79C: jge 00DFC7BEh
  loc_00DFC79E: push 000002B0h
  loc_00DFC7A3: push 006DCB00h
  loc_00DFC7A8: mov eax, Me
  loc_00DFC7AB: push eax
  loc_00DFC7AC: mov ecx, var_70
  loc_00DFC7AF: push ecx
  loc_00DFC7B0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC7B6: mov var_B0, eax
  loc_00DFC7BC: jmp 00DFC7C8h
  loc_00DFC7BE: mov var_B0, 00000000h
  loc_00DFC7C8: mov edx, var_28
  loc_00DFC7CB: mov var_74, edx
  loc_00DFC7CE: lea eax, var_2C
  loc_00DFC7D1: push eax
  loc_00DFC7D2: mov ecx, var_74
  loc_00DFC7D5: mov edx, [ecx]
  loc_00DFC7D7: mov eax, var_74
  loc_00DFC7DA: push eax
  loc_00DFC7DB: call [edx+00000024h]
  loc_00DFC7DE: fnclex
  loc_00DFC7E0: mov var_78, eax
  loc_00DFC7E3: cmp var_78, 00000000h
  loc_00DFC7E7: jge 00DFC806h
  loc_00DFC7E9: push 00000024h
  loc_00DFC7EB: push 006DEA70h
  loc_00DFC7F0: mov ecx, var_74
  loc_00DFC7F3: push ecx
  loc_00DFC7F4: mov edx, var_78
  loc_00DFC7F7: push edx
  loc_00DFC7F8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC7FE: mov var_B4, eax
  loc_00DFC804: jmp 00DFC810h
  loc_00DFC806: mov var_B4, 00000000h
  loc_00DFC810: mov eax, var_2C
  loc_00DFC813: mov var_A0, eax
  loc_00DFC819: mov var_2C, 00000000h
  loc_00DFC820: mov ecx, var_A0
  loc_00DFC826: mov var_38, ecx
  loc_00DFC829: mov var_40, 00000009h
  loc_00DFC830: lea edx, var_50
  loc_00DFC833: push edx
  loc_00DFC834: mov eax, 00000010h
  loc_00DFC839: call 00402830h ; __vbaChkstk
  loc_00DFC83E: mov eax, esp
  loc_00DFC840: mov ecx, var_40
  loc_00DFC843: mov [eax], ecx
  loc_00DFC845: mov edx, var_3C
  loc_00DFC848: mov [eax+00000004h], edx
  loc_00DFC84B: mov ecx, var_38
  loc_00DFC84E: mov [eax+00000008h], ecx
  loc_00DFC851: mov edx, var_34
  loc_00DFC854: mov [eax+0000000Ch], edx
  loc_00DFC857: push 006DEBACh ; "Font"
  loc_00DFC85C: mov eax, var_84
  loc_00DFC862: mov ecx, [eax]
  loc_00DFC864: mov edx, var_84
  loc_00DFC86A: push edx
  loc_00DFC86B: call [ecx+0000001Ch]
  loc_00DFC86E: fnclex
  loc_00DFC870: mov var_7C, eax
  loc_00DFC873: cmp var_7C, 00000000h
  loc_00DFC877: jge 00DFC899h
  loc_00DFC879: push 0000001Ch
  loc_00DFC87B: push 006DEB78h
  loc_00DFC880: mov eax, var_84
  loc_00DFC886: push eax
  loc_00DFC887: mov ecx, var_7C
  loc_00DFC88A: push ecx
  loc_00DFC88B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC891: mov var_B8, eax
  loc_00DFC897: jmp 00DFC8A3h
  loc_00DFC899: mov var_B8, 00000000h
  loc_00DFC8A3: push 006DDAD4h
  loc_00DFC8A8: lea edx, var_50
  loc_00DFC8AB: push edx
  loc_00DFC8AC: call [00401120h] ; __vbaCastObjVar
  loc_00DFC8B2: push eax
  loc_00DFC8B3: lea eax, var_30
  loc_00DFC8B6: push eax
  loc_00DFC8B7: call [004010ACh] ; __vbaObjSet
  loc_00DFC8BD: push eax
  loc_00DFC8BE: mov ecx, Me
  loc_00DFC8C1: mov edx, [ecx]
  loc_00DFC8C3: mov eax, Me
  loc_00DFC8C6: push eax
  loc_00DFC8C7: call [edx+00000A34h]
  loc_00DFC8CD: fnclex
  loc_00DFC8CF: mov var_80, eax
  loc_00DFC8D2: cmp var_80, 00000000h
  loc_00DFC8D6: jge 00DFC8F8h
  loc_00DFC8D8: push 00000A34h
  loc_00DFC8DD: push 006DD090h
  loc_00DFC8E2: mov ecx, Me
  loc_00DFC8E5: push ecx
  loc_00DFC8E6: mov edx, var_80
  loc_00DFC8E9: push edx
  loc_00DFC8EA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC8F0: mov var_BC, eax
  loc_00DFC8F6: jmp 00DFC902h
  loc_00DFC8F8: mov var_BC, 00000000h
  loc_00DFC902: lea eax, var_30
  loc_00DFC905: push eax
  loc_00DFC906: lea ecx, var_28
  loc_00DFC909: push ecx
  loc_00DFC90A: push 00000002h
  loc_00DFC90C: call [00401048h] ; __vbaFreeObjList
  loc_00DFC912: add esp, 0000000Ch
  loc_00DFC915: lea edx, var_50
  loc_00DFC918: push edx
  loc_00DFC919: lea eax, var_40
  loc_00DFC91C: push eax
  loc_00DFC91D: push 00000002h
  loc_00DFC91F: call [00401038h] ; __vbaFreeVarList
  loc_00DFC925: add esp, 0000000Ch
  loc_00DFC928: mov var_4, 00000006h
  loc_00DFC92F: lea ecx, var_28
  loc_00DFC932: push ecx
  loc_00DFC933: mov edx, Me
  loc_00DFC936: mov eax, [edx]
  loc_00DFC938: mov ecx, Me
  loc_00DFC93B: push ecx
  loc_00DFC93C: call [eax+00000A2Ch]
  loc_00DFC942: fnclex
  loc_00DFC944: mov var_70, eax
  loc_00DFC947: cmp var_70, 00000000h
  loc_00DFC94B: jge 00DFC96Dh
  loc_00DFC94D: push 00000A2Ch
  loc_00DFC952: push 006DD090h
  loc_00DFC957: mov edx, Me
  loc_00DFC95A: push edx
  loc_00DFC95B: mov eax, var_70
  loc_00DFC95E: push eax
  loc_00DFC95F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC965: mov var_C0, eax
  loc_00DFC96B: jmp 00DFC977h
  loc_00DFC96D: mov var_C0, 00000000h
  loc_00DFC977: mov ecx, var_28
  loc_00DFC97A: mov var_A4, ecx
  loc_00DFC980: mov var_28, 00000000h
  loc_00DFC987: mov edx, var_A4
  loc_00DFC98D: push edx
  loc_00DFC98E: lea eax, var_2C
  loc_00DFC991: push eax
  loc_00DFC992: call [004010ACh] ; __vbaObjSet
  loc_00DFC998: push eax
  loc_00DFC999: mov ecx, Me
  loc_00DFC99C: mov edx, [ecx+00000010h]
  loc_00DFC99F: mov eax, Me
  loc_00DFC9A2: mov ecx, [eax+00000010h]
  loc_00DFC9A5: mov eax, [ecx]
  loc_00DFC9A7: push edx
  loc_00DFC9A8: call [eax+0000024Ch]
  loc_00DFC9AE: fnclex
  loc_00DFC9B0: mov var_74, eax
  loc_00DFC9B3: cmp var_74, 00000000h
  loc_00DFC9B7: jge 00DFC9DCh
  loc_00DFC9B9: push 0000024Ch
  loc_00DFC9BE: push 006DCB00h
  loc_00DFC9C3: mov ecx, Me
  loc_00DFC9C6: mov edx, [ecx+00000010h]
  loc_00DFC9C9: push edx
  loc_00DFC9CA: mov eax, var_74
  loc_00DFC9CD: push eax
  loc_00DFC9CE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFC9D4: mov var_C4, eax
  loc_00DFC9DA: jmp 00DFC9E6h
  loc_00DFC9DC: mov var_C4, 00000000h
  loc_00DFC9E6: lea ecx, var_2C
  loc_00DFC9E9: call [00401254h] ; __vbaFreeObj
  loc_00DFC9EF: mov var_4, 00000007h
  loc_00DFC9F6: push 0000000Fh
  loc_00DFC9F8: call 006DDDB4h ; GetSysColor(%x1v)
  loc_00DFC9FD: mov var_68, eax
  loc_00DFCA00: call [00401074h] ; __vbaSetSystemError
  loc_00DFCA06: mov ecx, var_68
  loc_00DFCA09: mov var_58, ecx
  loc_00DFCA0C: mov var_60, 00000003h
  loc_00DFCA13: lea edx, var_40
  loc_00DFCA16: push edx
  loc_00DFCA17: mov eax, 00000010h
  loc_00DFCA1C: call 00402830h ; __vbaChkstk
  loc_00DFCA21: mov eax, esp
  loc_00DFCA23: mov ecx, var_60
  loc_00DFCA26: mov [eax], ecx
  loc_00DFCA28: mov edx, var_5C
  loc_00DFCA2B: mov [eax+00000004h], edx
  loc_00DFCA2E: mov ecx, var_58
  loc_00DFCA31: mov [eax+00000008h], ecx
  loc_00DFCA34: mov edx, var_54
  loc_00DFCA37: mov [eax+0000000Ch], edx
  loc_00DFCA3A: push 006DEACCh ; "BackColor"
  loc_00DFCA3F: mov eax, var_84
  loc_00DFCA45: mov ecx, [eax]
  loc_00DFCA47: mov edx, var_84
  loc_00DFCA4D: push edx
  loc_00DFCA4E: call [ecx+0000001Ch]
  loc_00DFCA51: fnclex
  loc_00DFCA53: mov var_70, eax
  loc_00DFCA56: cmp var_70, 00000000h
  loc_00DFCA5A: jge 00DFCA7Ch
  loc_00DFCA5C: push 0000001Ch
  loc_00DFCA5E: push 006DEB78h
  loc_00DFCA63: mov eax, var_84
  loc_00DFCA69: push eax
  loc_00DFCA6A: mov ecx, var_70
  loc_00DFCA6D: push ecx
  loc_00DFCA6E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCA74: mov var_C8, eax
  loc_00DFCA7A: jmp 00DFCA86h
  loc_00DFCA7C: mov var_C8, 00000000h
  loc_00DFCA86: lea edx, var_40
  loc_00DFCA89: push edx
  loc_00DFCA8A: call [004011D0h] ; __vbaI4Var
  loc_00DFCA90: mov ecx, Me
  loc_00DFCA93: mov [ecx+00000088h], eax
  loc_00DFCA99: lea ecx, var_40
  loc_00DFCA9C: call [00401024h] ; __vbaFreeVar
  loc_00DFCAA2: mov var_4, 00000008h
  loc_00DFCAA9: mov var_58, FFFFFFFFh
  loc_00DFCAB0: mov var_60, 0000000Bh
  loc_00DFCAB7: lea edx, var_40
  loc_00DFCABA: push edx
  loc_00DFCABB: mov eax, 00000010h
  loc_00DFCAC0: call 00402830h ; __vbaChkstk
  loc_00DFCAC5: mov eax, esp
  loc_00DFCAC7: mov ecx, var_60
  loc_00DFCACA: mov [eax], ecx
  loc_00DFCACC: mov edx, var_5C
  loc_00DFCACF: mov [eax+00000004h], edx
  loc_00DFCAD2: mov ecx, var_58
  loc_00DFCAD5: mov [eax+00000008h], ecx
  loc_00DFCAD8: mov edx, var_54
  loc_00DFCADB: mov [eax+0000000Ch], edx
  loc_00DFCADE: push 006DEBBCh ; "Enabled"
  loc_00DFCAE3: mov eax, var_84
  loc_00DFCAE9: mov ecx, [eax]
  loc_00DFCAEB: mov edx, var_84
  loc_00DFCAF1: push edx
  loc_00DFCAF2: call [ecx+0000001Ch]
  loc_00DFCAF5: fnclex
  loc_00DFCAF7: mov var_70, eax
  loc_00DFCAFA: cmp var_70, 00000000h
  loc_00DFCAFE: jge 00DFCB20h
  loc_00DFCB00: push 0000001Ch
  loc_00DFCB02: push 006DEB78h
  loc_00DFCB07: mov eax, var_84
  loc_00DFCB0D: push eax
  loc_00DFCB0E: mov ecx, var_70
  loc_00DFCB11: push ecx
  loc_00DFCB12: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCB18: mov var_CC, eax
  loc_00DFCB1E: jmp 00DFCB2Ah
  loc_00DFCB20: mov var_CC, 00000000h
  loc_00DFCB2A: lea edx, var_40
  loc_00DFCB2D: push edx
  loc_00DFCB2E: call [004010CCh] ; __vbaBoolVar
  loc_00DFCB34: mov ecx, Me
  loc_00DFCB37: mov [ecx+0000007Ch], ax
  loc_00DFCB3B: lea ecx, var_40
  loc_00DFCB3E: call [00401024h] ; __vbaFreeVar
  loc_00DFCB44: mov var_4, 00000009h
  loc_00DFCB4B: mov var_58, 006DEBE4h ; "jcbutton"
  loc_00DFCB52: mov var_60, 00000008h
  loc_00DFCB59: lea edx, var_40
  loc_00DFCB5C: push edx
  loc_00DFCB5D: mov eax, 00000010h
  loc_00DFCB62: call 00402830h ; __vbaChkstk
  loc_00DFCB67: mov eax, esp
  loc_00DFCB69: mov ecx, var_60
  loc_00DFCB6C: mov [eax], ecx
  loc_00DFCB6E: mov edx, var_5C
  loc_00DFCB71: mov [eax+00000004h], edx
  loc_00DFCB74: mov ecx, var_58
  loc_00DFCB77: mov [eax+00000008h], ecx
  loc_00DFCB7A: mov edx, var_54
  loc_00DFCB7D: mov [eax+0000000Ch], edx
  loc_00DFCB80: push 006DEBD0h ; "Caption"
  loc_00DFCB85: mov eax, var_84
  loc_00DFCB8B: mov ecx, [eax]
  loc_00DFCB8D: mov edx, var_84
  loc_00DFCB93: push edx
  loc_00DFCB94: call [ecx+0000001Ch]
  loc_00DFCB97: fnclex
  loc_00DFCB99: mov var_70, eax
  loc_00DFCB9C: cmp var_70, 00000000h
  loc_00DFCBA0: jge 00DFCBC2h
  loc_00DFCBA2: push 0000001Ch
  loc_00DFCBA4: push 006DEB78h
  loc_00DFCBA9: mov eax, var_84
  loc_00DFCBAF: push eax
  loc_00DFCBB0: mov ecx, var_70
  loc_00DFCBB3: push ecx
  loc_00DFCBB4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCBBA: mov var_D0, eax
  loc_00DFCBC0: jmp 00DFCBCCh
  loc_00DFCBC2: mov var_D0, 00000000h
  loc_00DFCBCC: lea edx, var_40
  loc_00DFCBCF: push edx
  loc_00DFCBD0: call [00401034h] ; __vbaStrVarMove
  loc_00DFCBD6: mov edx, eax
  loc_00DFCBD8: lea ecx, var_24
  loc_00DFCBDB: call [00401228h] ; __vbaStrMove
  loc_00DFCBE1: mov edx, eax
  loc_00DFCBE3: mov ecx, Me
  loc_00DFCBE6: add ecx, 00000080h
  loc_00DFCBEC: call [004011B0h] ; __vbaStrCopy
  loc_00DFCBF2: lea ecx, var_24
  loc_00DFCBF5: call [00401258h] ; __vbaFreeStr
  loc_00DFCBFB: lea ecx, var_40
  loc_00DFCBFE: call [00401024h] ; __vbaFreeVar
  loc_00DFCC04: mov var_4, 0000000Ah
  loc_00DFCC0B: mov var_58, 00000000h
  loc_00DFCC12: mov var_60, 0000000Bh
  loc_00DFCC19: lea eax, var_40
  loc_00DFCC1C: push eax
  loc_00DFCC1D: mov eax, 00000010h
  loc_00DFCC22: call 00402830h ; __vbaChkstk
  loc_00DFCC27: mov ecx, esp
  loc_00DFCC29: mov edx, var_60
  loc_00DFCC2C: mov [ecx], edx
  loc_00DFCC2E: mov eax, var_5C
  loc_00DFCC31: mov [ecx+00000004h], eax
  loc_00DFCC34: mov edx, var_58
  loc_00DFCC37: mov [ecx+00000008h], edx
  loc_00DFCC3A: mov eax, var_54
  loc_00DFCC3D: mov [ecx+0000000Ch], eax
  loc_00DFCC40: push 006DEBFCh ; "Value"
  loc_00DFCC45: mov ecx, var_84
  loc_00DFCC4B: mov edx, [ecx]
  loc_00DFCC4D: mov eax, var_84
  loc_00DFCC53: push eax
  loc_00DFCC54: call [edx+0000001Ch]
  loc_00DFCC57: fnclex
  loc_00DFCC59: mov var_70, eax
  loc_00DFCC5C: cmp var_70, 00000000h
  loc_00DFCC60: jge 00DFCC82h
  loc_00DFCC62: push 0000001Ch
  loc_00DFCC64: push 006DEB78h
  loc_00DFCC69: mov ecx, var_84
  loc_00DFCC6F: push ecx
  loc_00DFCC70: mov edx, var_70
  loc_00DFCC73: push edx
  loc_00DFCC74: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCC7A: mov var_D4, eax
  loc_00DFCC80: jmp 00DFCC8Ch
  loc_00DFCC82: mov var_D4, 00000000h
  loc_00DFCC8C: lea eax, var_40
  loc_00DFCC8F: push eax
  loc_00DFCC90: call [004010CCh] ; __vbaBoolVar
  loc_00DFCC96: mov ecx, Me
  loc_00DFCC99: mov [ecx+0000006Ch], ax
  loc_00DFCC9D: lea ecx, var_40
  loc_00DFCCA0: call [00401024h] ; __vbaFreeVar
  loc_00DFCCA6: mov var_4, 0000000Bh
  loc_00DFCCAD: mov var_58, 00000000h
  loc_00DFCCB4: mov var_60, 00000002h
  loc_00DFCCBB: lea edx, var_40
  loc_00DFCCBE: push edx
  loc_00DFCCBF: mov eax, 00000010h
  loc_00DFCCC4: call 00402830h ; __vbaChkstk
  loc_00DFCCC9: mov eax, esp
  loc_00DFCCCB: mov ecx, var_60
  loc_00DFCCCE: mov [eax], ecx
  loc_00DFCCD0: mov edx, var_5C
  loc_00DFCCD3: mov [eax+00000004h], edx
  loc_00DFCCD6: mov ecx, var_58
  loc_00DFCCD9: mov [eax+00000008h], ecx
  loc_00DFCCDC: mov edx, var_54
  loc_00DFCCDF: mov [eax+0000000Ch], edx
  loc_00DFCCE2: push 006DEC0Ch ; "MousePointer"
  loc_00DFCCE7: mov eax, var_84
  loc_00DFCCED: mov ecx, [eax]
  loc_00DFCCEF: mov edx, var_84
  loc_00DFCCF5: push edx
  loc_00DFCCF6: call [ecx+0000001Ch]
  loc_00DFCCF9: fnclex
  loc_00DFCCFB: mov var_70, eax
  loc_00DFCCFE: cmp var_70, 00000000h
  loc_00DFCD02: jge 00DFCD24h
  loc_00DFCD04: push 0000001Ch
  loc_00DFCD06: push 006DEB78h
  loc_00DFCD0B: mov eax, var_84
  loc_00DFCD11: push eax
  loc_00DFCD12: mov ecx, var_70
  loc_00DFCD15: push ecx
  loc_00DFCD16: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCD1C: mov var_D8, eax
  loc_00DFCD22: jmp 00DFCD2Eh
  loc_00DFCD24: mov var_D8, 00000000h
  loc_00DFCD2E: lea edx, var_40
  loc_00DFCD31: push edx
  loc_00DFCD32: call [0040118Ch] ; __vbaI2Var
  loc_00DFCD38: push eax
  loc_00DFCD39: mov eax, Me
  loc_00DFCD3C: mov ecx, [eax+00000010h]
  loc_00DFCD3F: mov edx, Me
  loc_00DFCD42: mov eax, [edx+00000010h]
  loc_00DFCD45: mov edx, [eax]
  loc_00DFCD47: push ecx
  loc_00DFCD48: call [edx+000000A4h]
  loc_00DFCD4E: fnclex
  loc_00DFCD50: mov var_74, eax
  loc_00DFCD53: cmp var_74, 00000000h
  loc_00DFCD57: jge 00DFCD7Ch
  loc_00DFCD59: push 000000A4h
  loc_00DFCD5E: push 006DCB00h
  loc_00DFCD63: mov eax, Me
  loc_00DFCD66: mov ecx, [eax+00000010h]
  loc_00DFCD69: push ecx
  loc_00DFCD6A: mov edx, var_74
  loc_00DFCD6D: push edx
  loc_00DFCD6E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCD74: mov var_DC, eax
  loc_00DFCD7A: jmp 00DFCD86h
  loc_00DFCD7C: mov var_DC, 00000000h
  loc_00DFCD86: lea ecx, var_40
  loc_00DFCD89: call [00401024h] ; __vbaFreeVar
  loc_00DFCD8F: mov var_4, 0000000Ch
  loc_00DFCD96: mov var_58, 00000000h
  loc_00DFCD9D: mov var_60, 0000000Bh
  loc_00DFCDA4: lea eax, var_40
  loc_00DFCDA7: push eax
  loc_00DFCDA8: mov eax, 00000010h
  loc_00DFCDAD: call 00402830h ; __vbaChkstk
  loc_00DFCDB2: mov ecx, esp
  loc_00DFCDB4: mov edx, var_60
  loc_00DFCDB7: mov [ecx], edx
  loc_00DFCDB9: mov eax, var_5C
  loc_00DFCDBC: mov [ecx+00000004h], eax
  loc_00DFCDBF: mov edx, var_58
  loc_00DFCDC2: mov [ecx+00000008h], edx
  loc_00DFCDC5: mov eax, var_54
  loc_00DFCDC8: mov [ecx+0000000Ch], eax
  loc_00DFCDCB: push 006DEC2Ch ; "HandPointer"
  loc_00DFCDD0: mov ecx, var_84
  loc_00DFCDD6: mov edx, [ecx]
  loc_00DFCDD8: mov eax, var_84
  loc_00DFCDDE: push eax
  loc_00DFCDDF: call [edx+0000001Ch]
  loc_00DFCDE2: fnclex
  loc_00DFCDE4: mov var_70, eax
  loc_00DFCDE7: cmp var_70, 00000000h
  loc_00DFCDEB: jge 00DFCE0Dh
  loc_00DFCDED: push 0000001Ch
  loc_00DFCDEF: push 006DEB78h
  loc_00DFCDF4: mov ecx, var_84
  loc_00DFCDFA: push ecx
  loc_00DFCDFB: mov edx, var_70
  loc_00DFCDFE: push edx
  loc_00DFCDFF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCE05: mov var_E0, eax
  loc_00DFCE0B: jmp 00DFCE17h
  loc_00DFCE0D: mov var_E0, 00000000h
  loc_00DFCE17: lea eax, var_40
  loc_00DFCE1A: push eax
  loc_00DFCE1B: call [004010CCh] ; __vbaBoolVar
  loc_00DFCE21: mov ecx, Me
  loc_00DFCE24: mov [ecx+00000052h], ax
  loc_00DFCE28: lea ecx, var_40
  loc_00DFCE2B: call [00401024h] ; __vbaFreeVar
  loc_00DFCE31: mov var_4, 0000000Dh
  loc_00DFCE38: mov var_38, 00000000h
  loc_00DFCE3F: mov var_40, 00000009h
  loc_00DFCE46: lea edx, var_50
  loc_00DFCE49: push edx
  loc_00DFCE4A: mov eax, 00000010h
  loc_00DFCE4F: call 00402830h ; __vbaChkstk
  loc_00DFCE54: mov eax, esp
  loc_00DFCE56: mov ecx, var_40
  loc_00DFCE59: mov [eax], ecx
  loc_00DFCE5B: mov edx, var_3C
  loc_00DFCE5E: mov [eax+00000004h], edx
  loc_00DFCE61: mov ecx, var_38
  loc_00DFCE64: mov [eax+00000008h], ecx
  loc_00DFCE67: mov edx, var_34
  loc_00DFCE6A: mov [eax+0000000Ch], edx
  loc_00DFCE6D: push 006DEC54h ; "MouseIcon"
  loc_00DFCE72: mov eax, var_84
  loc_00DFCE78: mov ecx, [eax]
  loc_00DFCE7A: mov edx, var_84
  loc_00DFCE80: push edx
  loc_00DFCE81: call [ecx+0000001Ch]
  loc_00DFCE84: fnclex
  loc_00DFCE86: mov var_70, eax
  loc_00DFCE89: cmp var_70, 00000000h
  loc_00DFCE8D: jge 00DFCEAFh
  loc_00DFCE8F: push 0000001Ch
  loc_00DFCE91: push 006DEB78h
  loc_00DFCE96: mov eax, var_84
  loc_00DFCE9C: push eax
  loc_00DFCE9D: mov ecx, var_70
  loc_00DFCEA0: push ecx
  loc_00DFCEA1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCEA7: mov var_E4, eax
  loc_00DFCEAD: jmp 00DFCEB9h
  loc_00DFCEAF: mov var_E4, 00000000h
  loc_00DFCEB9: push 006DE764h
  loc_00DFCEBE: lea edx, var_50
  loc_00DFCEC1: push edx
  loc_00DFCEC2: call [00401120h] ; __vbaCastObjVar
  loc_00DFCEC8: push eax
  loc_00DFCEC9: lea eax, var_28
  loc_00DFCECC: push eax
  loc_00DFCECD: call [004010ACh] ; __vbaObjSet
  loc_00DFCED3: push eax
  loc_00DFCED4: mov ecx, Me
  loc_00DFCED7: mov edx, [ecx+00000010h]
  loc_00DFCEDA: mov eax, Me
  loc_00DFCEDD: mov ecx, [eax+00000010h]
  loc_00DFCEE0: mov eax, [ecx]
  loc_00DFCEE2: push edx
  loc_00DFCEE3: call [eax+00000224h]
  loc_00DFCEE9: fnclex
  loc_00DFCEEB: mov var_74, eax
  loc_00DFCEEE: cmp var_74, 00000000h
  loc_00DFCEF2: jge 00DFCF17h
  loc_00DFCEF4: push 00000224h
  loc_00DFCEF9: push 006DCB00h
  loc_00DFCEFE: mov ecx, Me
  loc_00DFCF01: mov edx, [ecx+00000010h]
  loc_00DFCF04: push edx
  loc_00DFCF05: mov eax, var_74
  loc_00DFCF08: push eax
  loc_00DFCF09: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCF0F: mov var_E8, eax
  loc_00DFCF15: jmp 00DFCF21h
  loc_00DFCF17: mov var_E8, 00000000h
  loc_00DFCF21: lea ecx, var_28
  loc_00DFCF24: call [00401254h] ; __vbaFreeObj
  loc_00DFCF2A: lea ecx, var_50
  loc_00DFCF2D: push ecx
  loc_00DFCF2E: lea edx, var_40
  loc_00DFCF31: push edx
  loc_00DFCF32: push 00000002h
  loc_00DFCF34: call [00401038h] ; __vbaFreeVarList
  loc_00DFCF3A: add esp, 0000000Ch
  loc_00DFCF3D: mov var_4, 0000000Eh
  loc_00DFCF44: mov var_38, 00000000h
  loc_00DFCF4B: mov var_40, 00000009h
  loc_00DFCF52: lea eax, var_50
  loc_00DFCF55: push eax
  loc_00DFCF56: mov eax, 00000010h
  loc_00DFCF5B: call 00402830h ; __vbaChkstk
  loc_00DFCF60: mov ecx, esp
  loc_00DFCF62: mov edx, var_40
  loc_00DFCF65: mov [ecx], edx
  loc_00DFCF67: mov eax, var_3C
  loc_00DFCF6A: mov [ecx+00000004h], eax
  loc_00DFCF6D: mov edx, var_38
  loc_00DFCF70: mov [ecx+00000008h], edx
  loc_00DFCF73: mov eax, var_34
  loc_00DFCF76: mov [ecx+0000000Ch], eax
  loc_00DFCF79: push 006DEC6Ch ; "PictureNormal"
  loc_00DFCF7E: mov ecx, var_84
  loc_00DFCF84: mov edx, [ecx]
  loc_00DFCF86: mov eax, var_84
  loc_00DFCF8C: push eax
  loc_00DFCF8D: call [edx+0000001Ch]
  loc_00DFCF90: fnclex
  loc_00DFCF92: mov var_70, eax
  loc_00DFCF95: cmp var_70, 00000000h
  loc_00DFCF99: jge 00DFCFBBh
  loc_00DFCF9B: push 0000001Ch
  loc_00DFCF9D: push 006DEB78h
  loc_00DFCFA2: mov ecx, var_84
  loc_00DFCFA8: push ecx
  loc_00DFCFA9: mov edx, var_70
  loc_00DFCFAC: push edx
  loc_00DFCFAD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFCFB3: mov var_EC, eax
  loc_00DFCFB9: jmp 00DFCFC5h
  loc_00DFCFBB: mov var_EC, 00000000h
  loc_00DFCFC5: push 006DE764h
  loc_00DFCFCA: lea eax, var_50
  loc_00DFCFCD: push eax
  loc_00DFCFCE: call [00401120h] ; __vbaCastObjVar
  loc_00DFCFD4: push eax
  loc_00DFCFD5: lea ecx, var_28
  loc_00DFCFD8: push ecx
  loc_00DFCFD9: call [004010ACh] ; __vbaObjSet
  loc_00DFCFDF: push eax
  loc_00DFCFE0: mov edx, Me
  loc_00DFCFE3: add edx, 0000014Ch
  loc_00DFCFE9: push edx
  loc_00DFCFEA: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFCFF0: lea ecx, var_28
  loc_00DFCFF3: call [00401254h] ; __vbaFreeObj
  loc_00DFCFF9: lea eax, var_50
  loc_00DFCFFC: push eax
  loc_00DFCFFD: lea ecx, var_40
  loc_00DFD000: push ecx
  loc_00DFD001: push 00000002h
  loc_00DFD003: call [00401038h] ; __vbaFreeVarList
  loc_00DFD009: add esp, 0000000Ch
  loc_00DFD00C: mov var_4, 0000000Fh
  loc_00DFD013: mov var_38, 00000000h
  loc_00DFD01A: mov var_40, 00000009h
  loc_00DFD021: lea edx, var_50
  loc_00DFD024: push edx
  loc_00DFD025: mov eax, 00000010h
  loc_00DFD02A: call 00402830h ; __vbaChkstk
  loc_00DFD02F: mov eax, esp
  loc_00DFD031: mov ecx, var_40
  loc_00DFD034: mov [eax], ecx
  loc_00DFD036: mov edx, var_3C
  loc_00DFD039: mov [eax+00000004h], edx
  loc_00DFD03C: mov ecx, var_38
  loc_00DFD03F: mov [eax+00000008h], ecx
  loc_00DFD042: mov edx, var_34
  loc_00DFD045: mov [eax+0000000Ch], edx
  loc_00DFD048: push 006DEC8Ch ; "PictureHot"
  loc_00DFD04D: mov eax, var_84
  loc_00DFD053: mov ecx, [eax]
  loc_00DFD055: mov edx, var_84
  loc_00DFD05B: push edx
  loc_00DFD05C: call [ecx+0000001Ch]
  loc_00DFD05F: fnclex
  loc_00DFD061: mov var_70, eax
  loc_00DFD064: cmp var_70, 00000000h
  loc_00DFD068: jge 00DFD08Ah
  loc_00DFD06A: push 0000001Ch
  loc_00DFD06C: push 006DEB78h
  loc_00DFD071: mov eax, var_84
  loc_00DFD077: push eax
  loc_00DFD078: mov ecx, var_70
  loc_00DFD07B: push ecx
  loc_00DFD07C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD082: mov var_F0, eax
  loc_00DFD088: jmp 00DFD094h
  loc_00DFD08A: mov var_F0, 00000000h
  loc_00DFD094: push 006DE764h
  loc_00DFD099: lea edx, var_50
  loc_00DFD09C: push edx
  loc_00DFD09D: call [00401120h] ; __vbaCastObjVar
  loc_00DFD0A3: push eax
  loc_00DFD0A4: lea eax, var_28
  loc_00DFD0A7: push eax
  loc_00DFD0A8: call [004010ACh] ; __vbaObjSet
  loc_00DFD0AE: push eax
  loc_00DFD0AF: mov ecx, Me
  loc_00DFD0B2: add ecx, 00000150h
  loc_00DFD0B8: push ecx
  loc_00DFD0B9: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFD0BF: lea ecx, var_28
  loc_00DFD0C2: call [00401254h] ; __vbaFreeObj
  loc_00DFD0C8: lea edx, var_50
  loc_00DFD0CB: push edx
  loc_00DFD0CC: lea eax, var_40
  loc_00DFD0CF: push eax
  loc_00DFD0D0: push 00000002h
  loc_00DFD0D2: call [00401038h] ; __vbaFreeVarList
  loc_00DFD0D8: add esp, 0000000Ch
  loc_00DFD0DB: mov var_4, 00000010h
  loc_00DFD0E2: mov var_38, 00000000h
  loc_00DFD0E9: mov var_40, 00000009h
  loc_00DFD0F0: lea ecx, var_50
  loc_00DFD0F3: push ecx
  loc_00DFD0F4: mov eax, 00000010h
  loc_00DFD0F9: call 00402830h ; __vbaChkstk
  loc_00DFD0FE: mov edx, esp
  loc_00DFD100: mov eax, var_40
  loc_00DFD103: mov [edx], eax
  loc_00DFD105: mov ecx, var_3C
  loc_00DFD108: mov [edx+00000004h], ecx
  loc_00DFD10B: mov eax, var_38
  loc_00DFD10E: mov [edx+00000008h], eax
  loc_00DFD111: mov ecx, var_34
  loc_00DFD114: mov [edx+0000000Ch], ecx
  loc_00DFD117: push 006DECA8h ; "PictureDown"
  loc_00DFD11C: mov edx, var_84
  loc_00DFD122: mov eax, [edx]
  loc_00DFD124: mov ecx, var_84
  loc_00DFD12A: push ecx
  loc_00DFD12B: call [eax+0000001Ch]
  loc_00DFD12E: fnclex
  loc_00DFD130: mov var_70, eax
  loc_00DFD133: cmp var_70, 00000000h
  loc_00DFD137: jge 00DFD159h
  loc_00DFD139: push 0000001Ch
  loc_00DFD13B: push 006DEB78h
  loc_00DFD140: mov edx, var_84
  loc_00DFD146: push edx
  loc_00DFD147: mov eax, var_70
  loc_00DFD14A: push eax
  loc_00DFD14B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD151: mov var_F4, eax
  loc_00DFD157: jmp 00DFD163h
  loc_00DFD159: mov var_F4, 00000000h
  loc_00DFD163: push 006DE764h
  loc_00DFD168: lea ecx, var_50
  loc_00DFD16B: push ecx
  loc_00DFD16C: call [00401120h] ; __vbaCastObjVar
  loc_00DFD172: push eax
  loc_00DFD173: lea edx, var_28
  loc_00DFD176: push edx
  loc_00DFD177: call [004010ACh] ; __vbaObjSet
  loc_00DFD17D: push eax
  loc_00DFD17E: mov eax, Me
  loc_00DFD181: add eax, 00000154h
  loc_00DFD186: push eax
  loc_00DFD187: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFD18D: lea ecx, var_28
  loc_00DFD190: call [00401254h] ; __vbaFreeObj
  loc_00DFD196: lea ecx, var_50
  loc_00DFD199: push ecx
  loc_00DFD19A: lea edx, var_40
  loc_00DFD19D: push edx
  loc_00DFD19E: push 00000002h
  loc_00DFD1A0: call [00401038h] ; __vbaFreeVarList
  loc_00DFD1A6: add esp, 0000000Ch
  loc_00DFD1A9: mov var_4, 00000011h
  loc_00DFD1B0: mov var_58, 00000000h
  loc_00DFD1B7: mov var_60, 0000000Bh
  loc_00DFD1BE: lea eax, var_40
  loc_00DFD1C1: push eax
  loc_00DFD1C2: mov eax, 00000010h
  loc_00DFD1C7: call 00402830h ; __vbaChkstk
  loc_00DFD1CC: mov ecx, esp
  loc_00DFD1CE: mov edx, var_60
  loc_00DFD1D1: mov [ecx], edx
  loc_00DFD1D3: mov eax, var_5C
  loc_00DFD1D6: mov [ecx+00000004h], eax
  loc_00DFD1D9: mov edx, var_58
  loc_00DFD1DC: mov [ecx+00000008h], edx
  loc_00DFD1DF: mov eax, var_54
  loc_00DFD1E2: mov [ecx+0000000Ch], eax
  loc_00DFD1E5: push 006DECC4h ; "PictureShadow"
  loc_00DFD1EA: mov ecx, var_84
  loc_00DFD1F0: mov edx, [ecx]
  loc_00DFD1F2: mov eax, var_84
  loc_00DFD1F8: push eax
  loc_00DFD1F9: call [edx+0000001Ch]
  loc_00DFD1FC: fnclex
  loc_00DFD1FE: mov var_70, eax
  loc_00DFD201: cmp var_70, 00000000h
  loc_00DFD205: jge 00DFD227h
  loc_00DFD207: push 0000001Ch
  loc_00DFD209: push 006DEB78h
  loc_00DFD20E: mov ecx, var_84
  loc_00DFD214: push ecx
  loc_00DFD215: mov edx, var_70
  loc_00DFD218: push edx
  loc_00DFD219: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD21F: mov var_F8, eax
  loc_00DFD225: jmp 00DFD231h
  loc_00DFD227: mov var_F8, 00000000h
  loc_00DFD231: lea eax, var_40
  loc_00DFD234: push eax
  loc_00DFD235: call [004010CCh] ; __vbaBoolVar
  loc_00DFD23B: mov ecx, Me
  loc_00DFD23E: mov [ecx+00000160h], ax
  loc_00DFD245: lea ecx, var_40
  loc_00DFD248: call [00401024h] ; __vbaFreeVar
  loc_00DFD24E: mov var_4, 00000012h
  loc_00DFD255: mov var_58, 000000FFh
  loc_00DFD25C: mov var_60, 00000002h
  loc_00DFD263: lea edx, var_40
  loc_00DFD266: push edx
  loc_00DFD267: mov eax, 00000010h
  loc_00DFD26C: call 00402830h ; __vbaChkstk
  loc_00DFD271: mov eax, esp
  loc_00DFD273: mov ecx, var_60
  loc_00DFD276: mov [eax], ecx
  loc_00DFD278: mov edx, var_5C
  loc_00DFD27B: mov [eax+00000004h], edx
  loc_00DFD27E: mov ecx, var_58
  loc_00DFD281: mov [eax+00000008h], ecx
  loc_00DFD284: mov edx, var_54
  loc_00DFD287: mov [eax+0000000Ch], edx
  loc_00DFD28A: push 006DECE4h ; "PictureOpacity"
  loc_00DFD28F: mov eax, var_84
  loc_00DFD295: mov ecx, [eax]
  loc_00DFD297: mov edx, var_84
  loc_00DFD29D: push edx
  loc_00DFD29E: call [ecx+0000001Ch]
  loc_00DFD2A1: fnclex
  loc_00DFD2A3: mov var_70, eax
  loc_00DFD2A6: cmp var_70, 00000000h
  loc_00DFD2AA: jge 00DFD2CCh
  loc_00DFD2AC: push 0000001Ch
  loc_00DFD2AE: push 006DEB78h
  loc_00DFD2B3: mov eax, var_84
  loc_00DFD2B9: push eax
  loc_00DFD2BA: mov ecx, var_70
  loc_00DFD2BD: push ecx
  loc_00DFD2BE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD2C4: mov var_FC, eax
  loc_00DFD2CA: jmp 00DFD2D6h
  loc_00DFD2CC: mov var_FC, 00000000h
  loc_00DFD2D6: lea edx, var_40
  loc_00DFD2D9: push edx
  loc_00DFD2DA: call [00401248h] ; __vbaUI1Var
  loc_00DFD2E0: mov ecx, Me
  loc_00DFD2E3: mov [ecx+00000158h], al
  loc_00DFD2E9: lea ecx, var_40
  loc_00DFD2EC: call [00401024h] ; __vbaFreeVar
  loc_00DFD2F2: mov var_4, 00000013h
  loc_00DFD2F9: mov var_58, 000000FFh
  loc_00DFD300: mov var_60, 00000002h
  loc_00DFD307: lea edx, var_40
  loc_00DFD30A: push edx
  loc_00DFD30B: mov eax, 00000010h
  loc_00DFD310: call 00402830h ; __vbaChkstk
  loc_00DFD315: mov eax, esp
  loc_00DFD317: mov ecx, var_60
  loc_00DFD31A: mov [eax], ecx
  loc_00DFD31C: mov edx, var_5C
  loc_00DFD31F: mov [eax+00000004h], edx
  loc_00DFD322: mov ecx, var_58
  loc_00DFD325: mov [eax+00000008h], ecx
  loc_00DFD328: mov edx, var_54
  loc_00DFD32B: mov [eax+0000000Ch], edx
  loc_00DFD32E: push 006DED08h ; "PictureOpacityOnOver"
  loc_00DFD333: mov eax, var_84
  loc_00DFD339: mov ecx, [eax]
  loc_00DFD33B: mov edx, var_84
  loc_00DFD341: push edx
  loc_00DFD342: call [ecx+0000001Ch]
  loc_00DFD345: fnclex
  loc_00DFD347: mov var_70, eax
  loc_00DFD34A: cmp var_70, 00000000h
  loc_00DFD34E: jge 00DFD370h
  loc_00DFD350: push 0000001Ch
  loc_00DFD352: push 006DEB78h
  loc_00DFD357: mov eax, var_84
  loc_00DFD35D: push eax
  loc_00DFD35E: mov ecx, var_70
  loc_00DFD361: push ecx
  loc_00DFD362: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD368: mov var_100, eax
  loc_00DFD36E: jmp 00DFD37Ah
  loc_00DFD370: mov var_100, 00000000h
  loc_00DFD37A: lea edx, var_40
  loc_00DFD37D: push edx
  loc_00DFD37E: call [00401248h] ; __vbaUI1Var
  loc_00DFD384: mov ecx, Me
  loc_00DFD387: mov [ecx+00000159h], al
  loc_00DFD38D: lea ecx, var_40
  loc_00DFD390: call [00401024h] ; __vbaFreeVar
  loc_00DFD396: mov var_4, 00000014h
  loc_00DFD39D: mov var_58, 00000000h
  loc_00DFD3A4: mov var_60, 00000003h
  loc_00DFD3AB: lea edx, var_40
  loc_00DFD3AE: push edx
  loc_00DFD3AF: mov eax, 00000010h
  loc_00DFD3B4: call 00402830h ; __vbaChkstk
  loc_00DFD3B9: mov eax, esp
  loc_00DFD3BB: mov ecx, var_60
  loc_00DFD3BE: mov [eax], ecx
  loc_00DFD3C0: mov edx, var_5C
  loc_00DFD3C3: mov [eax+00000004h], edx
  loc_00DFD3C6: mov ecx, var_58
  loc_00DFD3C9: mov [eax+00000008h], ecx
  loc_00DFD3CC: mov edx, var_54
  loc_00DFD3CF: mov [eax+0000000Ch], edx
  loc_00DFD3D2: push 006DED38h ; "DisabledPictureMode"
  loc_00DFD3D7: mov eax, var_84
  loc_00DFD3DD: mov ecx, [eax]
  loc_00DFD3DF: mov edx, var_84
  loc_00DFD3E5: push edx
  loc_00DFD3E6: call [ecx+0000001Ch]
  loc_00DFD3E9: fnclex
  loc_00DFD3EB: mov var_70, eax
  loc_00DFD3EE: cmp var_70, 00000000h
  loc_00DFD3F2: jge 00DFD414h
  loc_00DFD3F4: push 0000001Ch
  loc_00DFD3F6: push 006DEB78h
  loc_00DFD3FB: mov eax, var_84
  loc_00DFD401: push eax
  loc_00DFD402: mov ecx, var_70
  loc_00DFD405: push ecx
  loc_00DFD406: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD40C: mov var_104, eax
  loc_00DFD412: jmp 00DFD41Eh
  loc_00DFD414: mov var_104, 00000000h
  loc_00DFD41E: lea edx, var_40
  loc_00DFD421: push edx
  loc_00DFD422: call [004011D0h] ; __vbaI4Var
  loc_00DFD428: mov ecx, Me
  loc_00DFD42B: mov [ecx+0000015Ch], eax
  loc_00DFD431: lea ecx, var_40
  loc_00DFD434: call [00401024h] ; __vbaFreeVar
  loc_00DFD43A: mov var_4, 00000015h
  loc_00DFD441: mov var_58, 00000000h
  loc_00DFD448: mov var_60, 0000000Bh
  loc_00DFD44F: lea edx, var_40
  loc_00DFD452: push edx
  loc_00DFD453: mov eax, 00000010h
  loc_00DFD458: call 00402830h ; __vbaChkstk
  loc_00DFD45D: mov eax, esp
  loc_00DFD45F: mov ecx, var_60
  loc_00DFD462: mov [eax], ecx
  loc_00DFD464: mov edx, var_5C
  loc_00DFD467: mov [eax+00000004h], edx
  loc_00DFD46A: mov ecx, var_58
  loc_00DFD46D: mov [eax+00000008h], ecx
  loc_00DFD470: mov edx, var_54
  loc_00DFD473: mov [eax+0000000Ch], edx
  loc_00DFD476: push 006DED64h ; "PicturePushOnHover"
  loc_00DFD47B: mov eax, var_84
  loc_00DFD481: mov ecx, [eax]
  loc_00DFD483: mov edx, var_84
  loc_00DFD489: push edx
  loc_00DFD48A: call [ecx+0000001Ch]
  loc_00DFD48D: fnclex
  loc_00DFD48F: mov var_70, eax
  loc_00DFD492: cmp var_70, 00000000h
  loc_00DFD496: jge 00DFD4B8h
  loc_00DFD498: push 0000001Ch
  loc_00DFD49A: push 006DEB78h
  loc_00DFD49F: mov eax, var_84
  loc_00DFD4A5: push eax
  loc_00DFD4A6: mov ecx, var_70
  loc_00DFD4A9: push ecx
  loc_00DFD4AA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD4B0: mov var_108, eax
  loc_00DFD4B6: jmp 00DFD4C2h
  loc_00DFD4B8: mov var_108, 00000000h
  loc_00DFD4C2: lea edx, var_40
  loc_00DFD4C5: push edx
  loc_00DFD4C6: call [004010CCh] ; __vbaBoolVar
  loc_00DFD4CC: mov ecx, Me
  loc_00DFD4CF: mov [ecx+00000170h], ax
  loc_00DFD4D6: lea ecx, var_40
  loc_00DFD4D9: call [00401024h] ; __vbaFreeVar
  loc_00DFD4DF: mov var_4, 00000016h
  loc_00DFD4E6: mov var_58, 00000001h
  loc_00DFD4ED: mov var_60, 00000003h
  loc_00DFD4F4: lea edx, var_40
  loc_00DFD4F7: push edx
  loc_00DFD4F8: mov eax, 00000010h
  loc_00DFD4FD: call 00402830h ; __vbaChkstk
  loc_00DFD502: mov eax, esp
  loc_00DFD504: mov ecx, var_60
  loc_00DFD507: mov [eax], ecx
  loc_00DFD509: mov edx, var_5C
  loc_00DFD50C: mov [eax+00000004h], edx
  loc_00DFD50F: mov ecx, var_58
  loc_00DFD512: mov [eax+00000008h], ecx
  loc_00DFD515: mov edx, var_54
  loc_00DFD518: mov [eax+0000000Ch], edx
  loc_00DFD51B: push 006DED90h ; "PictureEffectOnOver"
  loc_00DFD520: mov eax, var_84
  loc_00DFD526: mov ecx, [eax]
  loc_00DFD528: mov edx, var_84
  loc_00DFD52E: push edx
  loc_00DFD52F: call [ecx+0000001Ch]
  loc_00DFD532: fnclex
  loc_00DFD534: mov var_70, eax
  loc_00DFD537: cmp var_70, 00000000h
  loc_00DFD53B: jge 00DFD55Dh
  loc_00DFD53D: push 0000001Ch
  loc_00DFD53F: push 006DEB78h
  loc_00DFD544: mov eax, var_84
  loc_00DFD54A: push eax
  loc_00DFD54B: mov ecx, var_70
  loc_00DFD54E: push ecx
  loc_00DFD54F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD555: mov var_10C, eax
  loc_00DFD55B: jmp 00DFD567h
  loc_00DFD55D: mov var_10C, 00000000h
  loc_00DFD567: lea edx, var_40
  loc_00DFD56A: push edx
  loc_00DFD56B: call [004011D0h] ; __vbaI4Var
  loc_00DFD571: mov ecx, Me
  loc_00DFD574: mov [ecx+00000168h], eax
  loc_00DFD57A: lea ecx, var_40
  loc_00DFD57D: call [00401024h] ; __vbaFreeVar
  loc_00DFD583: mov var_4, 00000017h
  loc_00DFD58A: mov var_58, 00000002h
  loc_00DFD591: mov var_60, 00000003h
  loc_00DFD598: lea edx, var_40
  loc_00DFD59B: push edx
  loc_00DFD59C: mov eax, 00000010h
  loc_00DFD5A1: call 00402830h ; __vbaChkstk
  loc_00DFD5A6: mov eax, esp
  loc_00DFD5A8: mov ecx, var_60
  loc_00DFD5AB: mov [eax], ecx
  loc_00DFD5AD: mov edx, var_5C
  loc_00DFD5B0: mov [eax+00000004h], edx
  loc_00DFD5B3: mov ecx, var_58
  loc_00DFD5B6: mov [eax+00000008h], ecx
  loc_00DFD5B9: mov edx, var_54
  loc_00DFD5BC: mov [eax+0000000Ch], edx
  loc_00DFD5BF: push 006DEDBCh ; "PictureEffectOnDown"
  loc_00DFD5C4: mov eax, var_84
  loc_00DFD5CA: mov ecx, [eax]
  loc_00DFD5CC: mov edx, var_84
  loc_00DFD5D2: push edx
  loc_00DFD5D3: call [ecx+0000001Ch]
  loc_00DFD5D6: fnclex
  loc_00DFD5D8: mov var_70, eax
  loc_00DFD5DB: cmp var_70, 00000000h
  loc_00DFD5DF: jge 00DFD601h
  loc_00DFD5E1: push 0000001Ch
  loc_00DFD5E3: push 006DEB78h
  loc_00DFD5E8: mov eax, var_84
  loc_00DFD5EE: push eax
  loc_00DFD5EF: mov ecx, var_70
  loc_00DFD5F2: push ecx
  loc_00DFD5F3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD5F9: mov var_110, eax
  loc_00DFD5FF: jmp 00DFD60Bh
  loc_00DFD601: mov var_110, 00000000h
  loc_00DFD60B: lea edx, var_40
  loc_00DFD60E: push edx
  loc_00DFD60F: call [004011D0h] ; __vbaI4Var
  loc_00DFD615: mov ecx, Me
  loc_00DFD618: mov [ecx+0000016Ch], eax
  loc_00DFD61E: lea ecx, var_40
  loc_00DFD621: call [00401024h] ; __vbaFreeVar
  loc_00DFD627: mov var_4, 00000018h
  loc_00DFD62E: mov var_58, 00E0E0E0h
  loc_00DFD635: mov var_60, 00000003h
  loc_00DFD63C: lea edx, var_40
  loc_00DFD63F: push edx
  loc_00DFD640: mov eax, 00000010h
  loc_00DFD645: call 00402830h ; __vbaChkstk
  loc_00DFD64A: mov eax, esp
  loc_00DFD64C: mov ecx, var_60
  loc_00DFD64F: mov [eax], ecx
  loc_00DFD651: mov edx, var_5C
  loc_00DFD654: mov [eax+00000004h], edx
  loc_00DFD657: mov ecx, var_58
  loc_00DFD65A: mov [eax+00000008h], ecx
  loc_00DFD65D: mov edx, var_54
  loc_00DFD660: mov [eax+0000000Ch], edx
  loc_00DFD663: push 006DEDE8h ; "MaskColor"
  loc_00DFD668: mov eax, var_84
  loc_00DFD66E: mov ecx, [eax]
  loc_00DFD670: mov edx, var_84
  loc_00DFD676: push edx
  loc_00DFD677: call [ecx+0000001Ch]
  loc_00DFD67A: fnclex
  loc_00DFD67C: mov var_70, eax
  loc_00DFD67F: cmp var_70, 00000000h
  loc_00DFD683: jge 00DFD6A5h
  loc_00DFD685: push 0000001Ch
  loc_00DFD687: push 006DEB78h
  loc_00DFD68C: mov eax, var_84
  loc_00DFD692: push eax
  loc_00DFD693: mov ecx, var_70
  loc_00DFD696: push ecx
  loc_00DFD697: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD69D: mov var_114, eax
  loc_00DFD6A3: jmp 00DFD6AFh
  loc_00DFD6A5: mov var_114, 00000000h
  loc_00DFD6AF: lea edx, var_40
  loc_00DFD6B2: push edx
  loc_00DFD6B3: call [004011D0h] ; __vbaI4Var
  loc_00DFD6B9: mov ecx, Me
  loc_00DFD6BC: mov [ecx+000000A0h], eax
  loc_00DFD6C2: lea ecx, var_40
  loc_00DFD6C5: call [00401024h] ; __vbaFreeVar
  loc_00DFD6CB: mov var_4, 00000019h
  loc_00DFD6D2: mov var_58, FFFFFFFFh
  loc_00DFD6D9: mov var_60, 0000000Bh
  loc_00DFD6E0: lea edx, var_40
  loc_00DFD6E3: push edx
  loc_00DFD6E4: mov eax, 00000010h
  loc_00DFD6E9: call 00402830h ; __vbaChkstk
  loc_00DFD6EE: mov eax, esp
  loc_00DFD6F0: mov ecx, var_60
  loc_00DFD6F3: mov [eax], ecx
  loc_00DFD6F5: mov edx, var_5C
  loc_00DFD6F8: mov [eax+00000004h], edx
  loc_00DFD6FB: mov ecx, var_58
  loc_00DFD6FE: mov [eax+00000008h], ecx
  loc_00DFD701: mov edx, var_54
  loc_00DFD704: mov [eax+0000000Ch], edx
  loc_00DFD707: push 006DEE00h ; "UseMaskColor"
  loc_00DFD70C: mov eax, var_84
  loc_00DFD712: mov ecx, [eax]
  loc_00DFD714: mov edx, var_84
  loc_00DFD71A: push edx
  loc_00DFD71B: call [ecx+0000001Ch]
  loc_00DFD71E: fnclex
  loc_00DFD720: mov var_70, eax
  loc_00DFD723: cmp var_70, 00000000h
  loc_00DFD727: jge 00DFD749h
  loc_00DFD729: push 0000001Ch
  loc_00DFD72B: push 006DEB78h
  loc_00DFD730: mov eax, var_84
  loc_00DFD736: push eax
  loc_00DFD737: mov ecx, var_70
  loc_00DFD73A: push ecx
  loc_00DFD73B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD741: mov var_118, eax
  loc_00DFD747: jmp 00DFD753h
  loc_00DFD749: mov var_118, 00000000h
  loc_00DFD753: lea edx, var_40
  loc_00DFD756: push edx
  loc_00DFD757: call [004010CCh] ; __vbaBoolVar
  loc_00DFD75D: mov ecx, Me
  loc_00DFD760: mov [ecx+0000009Ch], ax
  loc_00DFD767: lea ecx, var_40
  loc_00DFD76A: call [00401024h] ; __vbaFreeVar
  loc_00DFD770: mov var_4, 0000001Ah
  loc_00DFD777: mov var_58, 00000000h
  loc_00DFD77E: mov var_60, 00000003h
  loc_00DFD785: lea edx, var_40
  loc_00DFD788: push edx
  loc_00DFD789: mov eax, 00000010h
  loc_00DFD78E: call 00402830h ; __vbaChkstk
  loc_00DFD793: mov eax, esp
  loc_00DFD795: mov ecx, var_60
  loc_00DFD798: mov [eax], ecx
  loc_00DFD79A: mov edx, var_5C
  loc_00DFD79D: mov [eax+00000004h], edx
  loc_00DFD7A0: mov ecx, var_58
  loc_00DFD7A3: mov [eax+00000008h], ecx
  loc_00DFD7A6: mov edx, var_54
  loc_00DFD7A9: mov [eax+0000000Ch], edx
  loc_00DFD7AC: push 006DEE20h ; "CaptionEffects"
  loc_00DFD7B1: mov eax, var_84
  loc_00DFD7B7: mov ecx, [eax]
  loc_00DFD7B9: mov edx, var_84
  loc_00DFD7BF: push edx
  loc_00DFD7C0: call [ecx+0000001Ch]
  loc_00DFD7C3: fnclex
  loc_00DFD7C5: mov var_70, eax
  loc_00DFD7C8: cmp var_70, 00000000h
  loc_00DFD7CC: jge 00DFD7EEh
  loc_00DFD7CE: push 0000001Ch
  loc_00DFD7D0: push 006DEB78h
  loc_00DFD7D5: mov eax, var_84
  loc_00DFD7DB: push eax
  loc_00DFD7DC: mov ecx, var_70
  loc_00DFD7DF: push ecx
  loc_00DFD7E0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD7E6: mov var_11C, eax
  loc_00DFD7EC: jmp 00DFD7F8h
  loc_00DFD7EE: mov var_11C, 00000000h
  loc_00DFD7F8: lea edx, var_40
  loc_00DFD7FB: push edx
  loc_00DFD7FC: call [004011D0h] ; __vbaI4Var
  loc_00DFD802: mov ecx, Me
  loc_00DFD805: mov [ecx+00000068h], eax
  loc_00DFD808: lea ecx, var_40
  loc_00DFD80B: call [00401024h] ; __vbaFreeVar
  loc_00DFD811: mov var_4, 0000001Bh
  loc_00DFD818: mov var_58, 00000000h
  loc_00DFD81F: mov var_60, 00000003h
  loc_00DFD826: lea edx, var_40
  loc_00DFD829: push edx
  loc_00DFD82A: mov eax, 00000010h
  loc_00DFD82F: call 00402830h ; __vbaChkstk
  loc_00DFD834: mov eax, esp
  loc_00DFD836: mov ecx, var_60
  loc_00DFD839: mov [eax], ecx
  loc_00DFD83B: mov edx, var_5C
  loc_00DFD83E: mov [eax+00000004h], edx
  loc_00DFD841: mov ecx, var_58
  loc_00DFD844: mov [eax+00000008h], ecx
  loc_00DFD847: mov edx, var_54
  loc_00DFD84A: mov [eax+0000000Ch], edx
  loc_00DFD84D: push 006DEE44h ; "Mode"
  loc_00DFD852: mov eax, var_84
  loc_00DFD858: mov ecx, [eax]
  loc_00DFD85A: mov edx, var_84
  loc_00DFD860: push edx
  loc_00DFD861: call [ecx+0000001Ch]
  loc_00DFD864: fnclex
  loc_00DFD866: mov var_70, eax
  loc_00DFD869: cmp var_70, 00000000h
  loc_00DFD86D: jge 00DFD88Fh
  loc_00DFD86F: push 0000001Ch
  loc_00DFD871: push 006DEB78h
  loc_00DFD876: mov eax, var_84
  loc_00DFD87C: push eax
  loc_00DFD87D: mov ecx, var_70
  loc_00DFD880: push ecx
  loc_00DFD881: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD887: mov var_120, eax
  loc_00DFD88D: jmp 00DFD899h
  loc_00DFD88F: mov var_120, 00000000h
  loc_00DFD899: lea edx, var_40
  loc_00DFD89C: push edx
  loc_00DFD89D: call [004011D0h] ; __vbaI4Var
  loc_00DFD8A3: mov ecx, Me
  loc_00DFD8A6: mov [ecx+00000064h], eax
  loc_00DFD8A9: lea ecx, var_40
  loc_00DFD8AC: call [00401024h] ; __vbaFreeVar
  loc_00DFD8B2: mov var_4, 0000001Ch
  loc_00DFD8B9: mov var_58, 00000001h
  loc_00DFD8C0: mov var_60, 00000003h
  loc_00DFD8C7: lea edx, var_40
  loc_00DFD8CA: push edx
  loc_00DFD8CB: mov eax, 00000010h
  loc_00DFD8D0: call 00402830h ; __vbaChkstk
  loc_00DFD8D5: mov eax, esp
  loc_00DFD8D7: mov ecx, var_60
  loc_00DFD8DA: mov [eax], ecx
  loc_00DFD8DC: mov edx, var_5C
  loc_00DFD8DF: mov [eax+00000004h], edx
  loc_00DFD8E2: mov ecx, var_58
  loc_00DFD8E5: mov [eax+00000008h], ecx
  loc_00DFD8E8: mov edx, var_54
  loc_00DFD8EB: mov [eax+0000000Ch], edx
  loc_00DFD8EE: push 006DEE54h ; "PictureAlign"
  loc_00DFD8F3: mov eax, var_84
  loc_00DFD8F9: mov ecx, [eax]
  loc_00DFD8FB: mov edx, var_84
  loc_00DFD901: push edx
  loc_00DFD902: call [ecx+0000001Ch]
  loc_00DFD905: fnclex
  loc_00DFD907: mov var_70, eax
  loc_00DFD90A: cmp var_70, 00000000h
  loc_00DFD90E: jge 00DFD930h
  loc_00DFD910: push 0000001Ch
  loc_00DFD912: push 006DEB78h
  loc_00DFD917: mov eax, var_84
  loc_00DFD91D: push eax
  loc_00DFD91E: mov ecx, var_70
  loc_00DFD921: push ecx
  loc_00DFD922: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD928: mov var_124, eax
  loc_00DFD92E: jmp 00DFD93Ah
  loc_00DFD930: mov var_124, 00000000h
  loc_00DFD93A: lea edx, var_40
  loc_00DFD93D: push edx
  loc_00DFD93E: call [004011D0h] ; __vbaI4Var
  loc_00DFD944: mov ecx, Me
  loc_00DFD947: mov [ecx+00000164h], eax
  loc_00DFD94D: lea ecx, var_40
  loc_00DFD950: call [00401024h] ; __vbaFreeVar
  loc_00DFD956: mov var_4, 0000001Dh
  loc_00DFD95D: mov var_58, 00000001h
  loc_00DFD964: mov var_60, 00000003h
  loc_00DFD96B: lea edx, var_40
  loc_00DFD96E: push edx
  loc_00DFD96F: mov eax, 00000010h
  loc_00DFD974: call 00402830h ; __vbaChkstk
  loc_00DFD979: mov eax, esp
  loc_00DFD97B: mov ecx, var_60
  loc_00DFD97E: mov [eax], ecx
  loc_00DFD980: mov edx, var_5C
  loc_00DFD983: mov [eax+00000004h], edx
  loc_00DFD986: mov ecx, var_58
  loc_00DFD989: mov [eax+00000008h], ecx
  loc_00DFD98C: mov edx, var_54
  loc_00DFD98F: mov [eax+0000000Ch], edx
  loc_00DFD992: push 006DEE74h ; "CaptionAlign"
  loc_00DFD997: mov eax, var_84
  loc_00DFD99D: mov ecx, [eax]
  loc_00DFD99F: mov edx, var_84
  loc_00DFD9A5: push edx
  loc_00DFD9A6: call [ecx+0000001Ch]
  loc_00DFD9A9: fnclex
  loc_00DFD9AB: mov var_70, eax
  loc_00DFD9AE: cmp var_70, 00000000h
  loc_00DFD9B2: jge 00DFD9D4h
  loc_00DFD9B4: push 0000001Ch
  loc_00DFD9B6: push 006DEB78h
  loc_00DFD9BB: mov eax, var_84
  loc_00DFD9C1: push eax
  loc_00DFD9C2: mov ecx, var_70
  loc_00DFD9C5: push ecx
  loc_00DFD9C6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFD9CC: mov var_128, eax
  loc_00DFD9D2: jmp 00DFD9DEh
  loc_00DFD9D4: mov var_128, 00000000h
  loc_00DFD9DE: lea edx, var_40
  loc_00DFD9E1: push edx
  loc_00DFD9E2: call [004011D0h] ; __vbaI4Var
  loc_00DFD9E8: mov ecx, Me
  loc_00DFD9EB: mov [ecx+00000084h], eax
  loc_00DFD9F1: lea ecx, var_40
  loc_00DFD9F4: call [00401024h] ; __vbaFreeVar
  loc_00DFD9FA: mov var_4, 0000001Eh
  loc_00DFDA01: mov var_68, 00000000h
  loc_00DFDA08: lea edx, var_6C
  loc_00DFDA0B: push edx
  loc_00DFDA0C: lea eax, var_68
  loc_00DFDA0F: push eax
  loc_00DFDA10: push 80000012h
  loc_00DFDA15: mov ecx, Me
  loc_00DFDA18: mov edx, [ecx]
  loc_00DFDA1A: mov eax, Me
  loc_00DFDA1D: push eax
  loc_00DFDA1E: call [edx+0000090Ch]
  loc_00DFDA24: mov ecx, var_6C
  loc_00DFDA27: mov var_58, ecx
  loc_00DFDA2A: mov var_60, 00000003h
  loc_00DFDA31: lea edx, var_40
  loc_00DFDA34: push edx
  loc_00DFDA35: mov eax, 00000010h
  loc_00DFDA3A: call 00402830h ; __vbaChkstk
  loc_00DFDA3F: mov eax, esp
  loc_00DFDA41: mov ecx, var_60
  loc_00DFDA44: mov [eax], ecx
  loc_00DFDA46: mov edx, var_5C
  loc_00DFDA49: mov [eax+00000004h], edx
  loc_00DFDA4C: mov ecx, var_58
  loc_00DFDA4F: mov [eax+00000008h], ecx
  loc_00DFDA52: mov edx, var_54
  loc_00DFDA55: mov [eax+0000000Ch], edx
  loc_00DFDA58: push 006DEE94h ; "ForeColor"
  loc_00DFDA5D: mov eax, var_84
  loc_00DFDA63: mov ecx, [eax]
  loc_00DFDA65: mov edx, var_84
  loc_00DFDA6B: push edx
  loc_00DFDA6C: call [ecx+0000001Ch]
  loc_00DFDA6F: fnclex
  loc_00DFDA71: mov var_70, eax
  loc_00DFDA74: cmp var_70, 00000000h
  loc_00DFDA78: jge 00DFDA9Ah
  loc_00DFDA7A: push 0000001Ch
  loc_00DFDA7C: push 006DEB78h
  loc_00DFDA81: mov eax, var_84
  loc_00DFDA87: push eax
  loc_00DFDA88: mov ecx, var_70
  loc_00DFDA8B: push ecx
  loc_00DFDA8C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDA92: mov var_12C, eax
  loc_00DFDA98: jmp 00DFDAA4h
  loc_00DFDA9A: mov var_12C, 00000000h
  loc_00DFDAA4: lea edx, var_40
  loc_00DFDAA7: push edx
  loc_00DFDAA8: call [004011D0h] ; __vbaI4Var
  loc_00DFDAAE: mov ecx, Me
  loc_00DFDAB1: mov [ecx+00000090h], eax
  loc_00DFDAB7: lea ecx, var_40
  loc_00DFDABA: call [00401024h] ; __vbaFreeVar
  loc_00DFDAC0: mov var_4, 0000001Fh
  loc_00DFDAC7: mov var_68, 00000000h
  loc_00DFDACE: lea edx, var_6C
  loc_00DFDAD1: push edx
  loc_00DFDAD2: lea eax, var_68
  loc_00DFDAD5: push eax
  loc_00DFDAD6: push 80000012h
  loc_00DFDADB: mov ecx, Me
  loc_00DFDADE: mov edx, [ecx]
  loc_00DFDAE0: mov eax, Me
  loc_00DFDAE3: push eax
  loc_00DFDAE4: call [edx+0000090Ch]
  loc_00DFDAEA: mov ecx, var_6C
  loc_00DFDAED: mov var_58, ecx
  loc_00DFDAF0: mov var_60, 00000003h
  loc_00DFDAF7: lea edx, var_40
  loc_00DFDAFA: push edx
  loc_00DFDAFB: mov eax, 00000010h
  loc_00DFDB00: call 00402830h ; __vbaChkstk
  loc_00DFDB05: mov eax, esp
  loc_00DFDB07: mov ecx, var_60
  loc_00DFDB0A: mov [eax], ecx
  loc_00DFDB0C: mov edx, var_5C
  loc_00DFDB0F: mov [eax+00000004h], edx
  loc_00DFDB12: mov ecx, var_58
  loc_00DFDB15: mov [eax+00000008h], ecx
  loc_00DFDB18: mov edx, var_54
  loc_00DFDB1B: mov [eax+0000000Ch], edx
  loc_00DFDB1E: push 006DEEACh ; "ForeColorHover"
  loc_00DFDB23: mov eax, var_84
  loc_00DFDB29: mov ecx, [eax]
  loc_00DFDB2B: mov edx, var_84
  loc_00DFDB31: push edx
  loc_00DFDB32: call [ecx+0000001Ch]
  loc_00DFDB35: fnclex
  loc_00DFDB37: mov var_70, eax
  loc_00DFDB3A: cmp var_70, 00000000h
  loc_00DFDB3E: jge 00DFDB60h
  loc_00DFDB40: push 0000001Ch
  loc_00DFDB42: push 006DEB78h
  loc_00DFDB47: mov eax, var_84
  loc_00DFDB4D: push eax
  loc_00DFDB4E: mov ecx, var_70
  loc_00DFDB51: push ecx
  loc_00DFDB52: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDB58: mov var_130, eax
  loc_00DFDB5E: jmp 00DFDB6Ah
  loc_00DFDB60: mov var_130, 00000000h
  loc_00DFDB6A: lea edx, var_40
  loc_00DFDB6D: push edx
  loc_00DFDB6E: call [004011D0h] ; __vbaI4Var
  loc_00DFDB74: mov ecx, Me
  loc_00DFDB77: mov [ecx+00000094h], eax
  loc_00DFDB7D: lea ecx, var_40
  loc_00DFDB80: call [00401024h] ; __vbaFreeVar
  loc_00DFDB86: mov var_4, 00000020h
  loc_00DFDB8D: mov edx, Me
  loc_00DFDB90: mov eax, [edx+00000090h]
  loc_00DFDB96: push eax
  loc_00DFDB97: mov ecx, Me
  loc_00DFDB9A: mov edx, [ecx+00000010h]
  loc_00DFDB9D: mov eax, Me
  loc_00DFDBA0: mov ecx, [eax+00000010h]
  loc_00DFDBA3: mov eax, [ecx]
  loc_00DFDBA5: push edx
  loc_00DFDBA6: call [eax+0000006Ch]
  loc_00DFDBA9: fnclex
  loc_00DFDBAB: mov var_70, eax
  loc_00DFDBAE: cmp var_70, 00000000h
  loc_00DFDBB2: jge 00DFDBD4h
  loc_00DFDBB4: push 0000006Ch
  loc_00DFDBB6: push 006DCB00h
  loc_00DFDBBB: mov ecx, Me
  loc_00DFDBBE: mov edx, [ecx+00000010h]
  loc_00DFDBC1: push edx
  loc_00DFDBC2: mov eax, var_70
  loc_00DFDBC5: push eax
  loc_00DFDBC6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDBCC: mov var_134, eax
  loc_00DFDBD2: jmp 00DFDBDEh
  loc_00DFDBD4: mov var_134, 00000000h
  loc_00DFDBDE: mov var_4, 00000021h
  loc_00DFDBE5: mov var_58, 00000000h
  loc_00DFDBEC: mov var_60, 0000000Bh
  loc_00DFDBF3: lea ecx, var_40
  loc_00DFDBF6: push ecx
  loc_00DFDBF7: mov eax, 00000010h
  loc_00DFDBFC: call 00402830h ; __vbaChkstk
  loc_00DFDC01: mov edx, esp
  loc_00DFDC03: mov eax, var_60
  loc_00DFDC06: mov [edx], eax
  loc_00DFDC08: mov ecx, var_5C
  loc_00DFDC0B: mov [edx+00000004h], ecx
  loc_00DFDC0E: mov eax, var_58
  loc_00DFDC11: mov [edx+00000008h], eax
  loc_00DFDC14: mov ecx, var_54
  loc_00DFDC17: mov [edx+0000000Ch], ecx
  loc_00DFDC1A: push 006DEED0h ; "DropDownSeparator"
  loc_00DFDC1F: mov edx, var_84
  loc_00DFDC25: mov eax, [edx]
  loc_00DFDC27: mov ecx, var_84
  loc_00DFDC2D: push ecx
  loc_00DFDC2E: call [eax+0000001Ch]
  loc_00DFDC31: fnclex
  loc_00DFDC33: mov var_70, eax
  loc_00DFDC36: cmp var_70, 00000000h
  loc_00DFDC3A: jge 00DFDC5Ch
  loc_00DFDC3C: push 0000001Ch
  loc_00DFDC3E: push 006DEB78h
  loc_00DFDC43: mov edx, var_84
  loc_00DFDC49: push edx
  loc_00DFDC4A: mov eax, var_70
  loc_00DFDC4D: push eax
  loc_00DFDC4E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDC54: mov var_138, eax
  loc_00DFDC5A: jmp 00DFDC66h
  loc_00DFDC5C: mov var_138, 00000000h
  loc_00DFDC66: lea ecx, var_40
  loc_00DFDC69: push ecx
  loc_00DFDC6A: call [004010CCh] ; __vbaBoolVar
  loc_00DFDC70: mov edx, Me
  loc_00DFDC73: mov [edx+00000060h], ax
  loc_00DFDC77: lea ecx, var_40
  loc_00DFDC7A: call [00401024h] ; __vbaFreeVar
  loc_00DFDC80: mov var_4, 00000022h
  loc_00DFDC87: mov var_58, 00000000h
  loc_00DFDC8E: mov var_60, 00000008h
  loc_00DFDC95: lea eax, var_40
  loc_00DFDC98: push eax
  loc_00DFDC99: mov eax, 00000010h
  loc_00DFDC9E: call 00402830h ; __vbaChkstk
  loc_00DFDCA3: mov ecx, esp
  loc_00DFDCA5: mov edx, var_60
  loc_00DFDCA8: mov [ecx], edx
  loc_00DFDCAA: mov eax, var_5C
  loc_00DFDCAD: mov [ecx+00000004h], eax
  loc_00DFDCB0: mov edx, var_58
  loc_00DFDCB3: mov [ecx+00000008h], edx
  loc_00DFDCB6: mov eax, var_54
  loc_00DFDCB9: mov [ecx+0000000Ch], eax
  loc_00DFDCBC: push 006DEEF8h ; "TooltipTitle"
  loc_00DFDCC1: mov ecx, var_84
  loc_00DFDCC7: mov edx, [ecx]
  loc_00DFDCC9: mov eax, var_84
  loc_00DFDCCF: push eax
  loc_00DFDCD0: call [edx+0000001Ch]
  loc_00DFDCD3: fnclex
  loc_00DFDCD5: mov var_70, eax
  loc_00DFDCD8: cmp var_70, 00000000h
  loc_00DFDCDC: jge 00DFDCFEh
  loc_00DFDCDE: push 0000001Ch
  loc_00DFDCE0: push 006DEB78h
  loc_00DFDCE5: mov ecx, var_84
  loc_00DFDCEB: push ecx
  loc_00DFDCEC: mov edx, var_70
  loc_00DFDCEF: push edx
  loc_00DFDCF0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDCF6: mov var_13C, eax
  loc_00DFDCFC: jmp 00DFDD08h
  loc_00DFDCFE: mov var_13C, 00000000h
  loc_00DFDD08: lea eax, var_40
  loc_00DFDD0B: push eax
  loc_00DFDD0C: call [00401034h] ; __vbaStrVarMove
  loc_00DFDD12: mov edx, eax
  loc_00DFDD14: lea ecx, var_24
  loc_00DFDD17: call [00401228h] ; __vbaStrMove
  loc_00DFDD1D: mov edx, eax
  loc_00DFDD1F: mov ecx, Me
  loc_00DFDD22: add ecx, 000000FCh
  loc_00DFDD28: call [004011B0h] ; __vbaStrCopy
  loc_00DFDD2E: lea ecx, var_24
  loc_00DFDD31: call [00401258h] ; __vbaFreeStr
  loc_00DFDD37: lea ecx, var_40
  loc_00DFDD3A: call [00401024h] ; __vbaFreeVar
  loc_00DFDD40: mov var_4, 00000023h
  loc_00DFDD47: mov var_58, 00000000h
  loc_00DFDD4E: mov var_60, 00000008h
  loc_00DFDD55: lea ecx, var_40
  loc_00DFDD58: push ecx
  loc_00DFDD59: mov eax, 00000010h
  loc_00DFDD5E: call 00402830h ; __vbaChkstk
  loc_00DFDD63: mov edx, esp
  loc_00DFDD65: mov eax, var_60
  loc_00DFDD68: mov [edx], eax
  loc_00DFDD6A: mov ecx, var_5C
  loc_00DFDD6D: mov [edx+00000004h], ecx
  loc_00DFDD70: mov eax, var_58
  loc_00DFDD73: mov [edx+00000008h], eax
  loc_00DFDD76: mov ecx, var_54
  loc_00DFDD79: mov [edx+0000000Ch], ecx
  loc_00DFDD7C: push 006DEF18h ; "ToolTip"
  loc_00DFDD81: mov edx, var_84
  loc_00DFDD87: mov eax, [edx]
  loc_00DFDD89: mov ecx, var_84
  loc_00DFDD8F: push ecx
  loc_00DFDD90: call [eax+0000001Ch]
  loc_00DFDD93: fnclex
  loc_00DFDD95: mov var_70, eax
  loc_00DFDD98: cmp var_70, 00000000h
  loc_00DFDD9C: jge 00DFDDBEh
  loc_00DFDD9E: push 0000001Ch
  loc_00DFDDA0: push 006DEB78h
  loc_00DFDDA5: mov edx, var_84
  loc_00DFDDAB: push edx
  loc_00DFDDAC: mov eax, var_70
  loc_00DFDDAF: push eax
  loc_00DFDDB0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDDB6: mov var_140, eax
  loc_00DFDDBC: jmp 00DFDDC8h
  loc_00DFDDBE: mov var_140, 00000000h
  loc_00DFDDC8: lea ecx, var_40
  loc_00DFDDCB: push ecx
  loc_00DFDDCC: call [00401034h] ; __vbaStrVarMove
  loc_00DFDDD2: mov edx, eax
  loc_00DFDDD4: lea ecx, var_24
  loc_00DFDDD7: call [00401228h] ; __vbaStrMove
  loc_00DFDDDD: mov edx, eax
  loc_00DFDDDF: mov ecx, Me
  loc_00DFDDE2: add ecx, 000000F8h
  loc_00DFDDE8: call [004011B0h] ; __vbaStrCopy
  loc_00DFDDEE: lea ecx, var_24
  loc_00DFDDF1: call [00401258h] ; __vbaFreeStr
  loc_00DFDDF7: lea ecx, var_40
  loc_00DFDDFA: call [00401024h] ; __vbaFreeVar
  loc_00DFDE00: mov var_4, 00000024h
  loc_00DFDE07: mov var_58, 00000000h
  loc_00DFDE0E: mov var_60, 00000003h
  loc_00DFDE15: lea edx, var_40
  loc_00DFDE18: push edx
  loc_00DFDE19: mov eax, 00000010h
  loc_00DFDE1E: call 00402830h ; __vbaChkstk
  loc_00DFDE23: mov eax, esp
  loc_00DFDE25: mov ecx, var_60
  loc_00DFDE28: mov [eax], ecx
  loc_00DFDE2A: mov edx, var_5C
  loc_00DFDE2D: mov [eax+00000004h], edx
  loc_00DFDE30: mov ecx, var_58
  loc_00DFDE33: mov [eax+00000008h], ecx
  loc_00DFDE36: mov edx, var_54
  loc_00DFDE39: mov [eax+0000000Ch], edx
  loc_00DFDE3C: push 006DEF2Ch ; "TooltipIcon"
  loc_00DFDE41: mov eax, var_84
  loc_00DFDE47: mov ecx, [eax]
  loc_00DFDE49: mov edx, var_84
  loc_00DFDE4F: push edx
  loc_00DFDE50: call [ecx+0000001Ch]
  loc_00DFDE53: fnclex
  loc_00DFDE55: mov var_70, eax
  loc_00DFDE58: cmp var_70, 00000000h
  loc_00DFDE5C: jge 00DFDE7Eh
  loc_00DFDE5E: push 0000001Ch
  loc_00DFDE60: push 006DEB78h
  loc_00DFDE65: mov eax, var_84
  loc_00DFDE6B: push eax
  loc_00DFDE6C: mov ecx, var_70
  loc_00DFDE6F: push ecx
  loc_00DFDE70: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDE76: mov var_144, eax
  loc_00DFDE7C: jmp 00DFDE88h
  loc_00DFDE7E: mov var_144, 00000000h
  loc_00DFDE88: lea edx, var_40
  loc_00DFDE8B: push edx
  loc_00DFDE8C: call [004011D0h] ; __vbaI4Var
  loc_00DFDE92: mov ecx, Me
  loc_00DFDE95: mov [ecx+00000100h], eax
  loc_00DFDE9B: lea ecx, var_40
  loc_00DFDE9E: call [00401024h] ; __vbaFreeVar
  loc_00DFDEA4: mov var_4, 00000025h
  loc_00DFDEAB: mov var_58, 00000000h
  loc_00DFDEB2: mov var_60, 00000003h
  loc_00DFDEB9: lea edx, var_40
  loc_00DFDEBC: push edx
  loc_00DFDEBD: mov eax, 00000010h
  loc_00DFDEC2: call 00402830h ; __vbaChkstk
  loc_00DFDEC7: mov eax, esp
  loc_00DFDEC9: mov ecx, var_60
  loc_00DFDECC: mov [eax], ecx
  loc_00DFDECE: mov edx, var_5C
  loc_00DFDED1: mov [eax+00000004h], edx
  loc_00DFDED4: mov ecx, var_58
  loc_00DFDED7: mov [eax+00000008h], ecx
  loc_00DFDEDA: mov edx, var_54
  loc_00DFDEDD: mov [eax+0000000Ch], edx
  loc_00DFDEE0: push 006DEF48h ; "TooltipType"
  loc_00DFDEE5: mov eax, var_84
  loc_00DFDEEB: mov ecx, [eax]
  loc_00DFDEED: mov edx, var_84
  loc_00DFDEF3: push edx
  loc_00DFDEF4: call [ecx+0000001Ch]
  loc_00DFDEF7: fnclex
  loc_00DFDEF9: mov var_70, eax
  loc_00DFDEFC: cmp var_70, 00000000h
  loc_00DFDF00: jge 00DFDF22h
  loc_00DFDF02: push 0000001Ch
  loc_00DFDF04: push 006DEB78h
  loc_00DFDF09: mov eax, var_84
  loc_00DFDF0F: push eax
  loc_00DFDF10: mov ecx, var_70
  loc_00DFDF13: push ecx
  loc_00DFDF14: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDF1A: mov var_148, eax
  loc_00DFDF20: jmp 00DFDF2Ch
  loc_00DFDF22: mov var_148, 00000000h
  loc_00DFDF2C: lea edx, var_40
  loc_00DFDF2F: push edx
  loc_00DFDF30: call [004011D0h] ; __vbaI4Var
  loc_00DFDF36: mov ecx, Me
  loc_00DFDF39: mov [ecx+00000104h], eax
  loc_00DFDF3F: lea ecx, var_40
  loc_00DFDF42: call [00401024h] ; __vbaFreeVar
  loc_00DFDF48: mov var_4, 00000026h
  loc_00DFDF4F: mov var_68, 00000000h
  loc_00DFDF56: lea edx, var_6C
  loc_00DFDF59: push edx
  loc_00DFDF5A: lea eax, var_68
  loc_00DFDF5D: push eax
  loc_00DFDF5E: push 80000018h
  loc_00DFDF63: mov ecx, Me
  loc_00DFDF66: mov edx, [ecx]
  loc_00DFDF68: mov eax, Me
  loc_00DFDF6B: push eax
  loc_00DFDF6C: call [edx+0000090Ch]
  loc_00DFDF72: mov ecx, var_6C
  loc_00DFDF75: mov var_58, ecx
  loc_00DFDF78: mov var_60, 00000003h
  loc_00DFDF7F: lea edx, var_40
  loc_00DFDF82: push edx
  loc_00DFDF83: mov eax, 00000010h
  loc_00DFDF88: call 00402830h ; __vbaChkstk
  loc_00DFDF8D: mov eax, esp
  loc_00DFDF8F: mov ecx, var_60
  loc_00DFDF92: mov [eax], ecx
  loc_00DFDF94: mov edx, var_5C
  loc_00DFDF97: mov [eax+00000004h], edx
  loc_00DFDF9A: mov ecx, var_58
  loc_00DFDF9D: mov [eax+00000008h], ecx
  loc_00DFDFA0: mov edx, var_54
  loc_00DFDFA3: mov [eax+0000000Ch], edx
  loc_00DFDFA6: push 006DEF64h ; "TooltipBackColor"
  loc_00DFDFAB: mov eax, var_84
  loc_00DFDFB1: mov ecx, [eax]
  loc_00DFDFB3: mov edx, var_84
  loc_00DFDFB9: push edx
  loc_00DFDFBA: call [ecx+0000001Ch]
  loc_00DFDFBD: fnclex
  loc_00DFDFBF: mov var_70, eax
  loc_00DFDFC2: cmp var_70, 00000000h
  loc_00DFDFC6: jge 00DFDFE8h
  loc_00DFDFC8: push 0000001Ch
  loc_00DFDFCA: push 006DEB78h
  loc_00DFDFCF: mov eax, var_84
  loc_00DFDFD5: push eax
  loc_00DFDFD6: mov ecx, var_70
  loc_00DFDFD9: push ecx
  loc_00DFDFDA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFDFE0: mov var_14C, eax
  loc_00DFDFE6: jmp 00DFDFF2h
  loc_00DFDFE8: mov var_14C, 00000000h
  loc_00DFDFF2: lea edx, var_40
  loc_00DFDFF5: push edx
  loc_00DFDFF6: call [004011D0h] ; __vbaI4Var
  loc_00DFDFFC: mov ecx, Me
  loc_00DFDFFF: mov [ecx+00000108h], eax
  loc_00DFE005: lea ecx, var_40
  loc_00DFE008: call [00401024h] ; __vbaFreeVar
  loc_00DFE00E: mov var_4, 00000027h
  loc_00DFE015: mov var_58, 00000000h
  loc_00DFE01C: mov var_60, 0000000Bh
  loc_00DFE023: lea edx, var_40
  loc_00DFE026: push edx
  loc_00DFE027: mov eax, 00000010h
  loc_00DFE02C: call 00402830h ; __vbaChkstk
  loc_00DFE031: mov eax, esp
  loc_00DFE033: mov ecx, var_60
  loc_00DFE036: mov [eax], ecx
  loc_00DFE038: mov edx, var_5C
  loc_00DFE03B: mov [eax+00000004h], edx
  loc_00DFE03E: mov ecx, var_58
  loc_00DFE041: mov [eax+00000008h], ecx
  loc_00DFE044: mov edx, var_54
  loc_00DFE047: mov [eax+0000000Ch], edx
  loc_00DFE04A: push 006DEF8Ch ; "RightToLeft"
  loc_00DFE04F: mov eax, var_84
  loc_00DFE055: mov ecx, [eax]
  loc_00DFE057: mov edx, var_84
  loc_00DFE05D: push edx
  loc_00DFE05E: call [ecx+0000001Ch]
  loc_00DFE061: fnclex
  loc_00DFE063: mov var_70, eax
  loc_00DFE066: cmp var_70, 00000000h
  loc_00DFE06A: jge 00DFE08Ch
  loc_00DFE06C: push 0000001Ch
  loc_00DFE06E: push 006DEB78h
  loc_00DFE073: mov eax, var_84
  loc_00DFE079: push eax
  loc_00DFE07A: mov ecx, var_70
  loc_00DFE07D: push ecx
  loc_00DFE07E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE084: mov var_150, eax
  loc_00DFE08A: jmp 00DFE096h
  loc_00DFE08C: mov var_150, 00000000h
  loc_00DFE096: lea edx, var_40
  loc_00DFE099: push edx
  loc_00DFE09A: call [004010CCh] ; __vbaBoolVar
  loc_00DFE0A0: mov ecx, Me
  loc_00DFE0A3: mov [ecx+00000138h], ax
  loc_00DFE0AA: lea ecx, var_40
  loc_00DFE0AD: call [00401024h] ; __vbaFreeVar
  loc_00DFE0B3: mov var_4, 00000028h
  loc_00DFE0BA: mov var_58, 00000000h
  loc_00DFE0C1: mov var_60, 0000000Bh
  loc_00DFE0C8: lea edx, var_40
  loc_00DFE0CB: push edx
  loc_00DFE0CC: mov eax, 00000010h
  loc_00DFE0D1: call 00402830h ; __vbaChkstk
  loc_00DFE0D6: mov eax, esp
  loc_00DFE0D8: mov ecx, var_60
  loc_00DFE0DB: mov [eax], ecx
  loc_00DFE0DD: mov edx, var_5C
  loc_00DFE0E0: mov [eax+00000004h], edx
  loc_00DFE0E3: mov ecx, var_58
  loc_00DFE0E6: mov [eax+00000008h], ecx
  loc_00DFE0E9: mov edx, var_54
  loc_00DFE0EC: mov [eax+0000000Ch], edx
  loc_00DFE0EF: push 006DEF8Ch ; "RightToLeft"
  loc_00DFE0F4: mov eax, var_84
  loc_00DFE0FA: mov ecx, [eax]
  loc_00DFE0FC: mov edx, var_84
  loc_00DFE102: push edx
  loc_00DFE103: call [ecx+0000001Ch]
  loc_00DFE106: fnclex
  loc_00DFE108: mov var_70, eax
  loc_00DFE10B: cmp var_70, 00000000h
  loc_00DFE10F: jge 00DFE131h
  loc_00DFE111: push 0000001Ch
  loc_00DFE113: push 006DEB78h
  loc_00DFE118: mov eax, var_84
  loc_00DFE11E: push eax
  loc_00DFE11F: mov ecx, var_70
  loc_00DFE122: push ecx
  loc_00DFE123: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE129: mov var_154, eax
  loc_00DFE12F: jmp 00DFE13Bh
  loc_00DFE131: mov var_154, 00000000h
  loc_00DFE13B: lea edx, var_40
  loc_00DFE13E: push edx
  loc_00DFE13F: call [004010CCh] ; __vbaBoolVar
  loc_00DFE145: mov ecx, Me
  loc_00DFE148: mov [ecx+0000010Eh], ax
  loc_00DFE14F: lea ecx, var_40
  loc_00DFE152: call [00401024h] ; __vbaFreeVar
  loc_00DFE158: mov var_4, 00000029h
  loc_00DFE15F: mov var_58, 00000000h
  loc_00DFE166: mov var_60, 00000003h
  loc_00DFE16D: lea edx, var_40
  loc_00DFE170: push edx
  loc_00DFE171: mov eax, 00000010h
  loc_00DFE176: call 00402830h ; __vbaChkstk
  loc_00DFE17B: mov eax, esp
  loc_00DFE17D: mov ecx, var_60
  loc_00DFE180: mov [eax], ecx
  loc_00DFE182: mov edx, var_5C
  loc_00DFE185: mov [eax+00000004h], edx
  loc_00DFE188: mov ecx, var_58
  loc_00DFE18B: mov [eax+00000008h], ecx
  loc_00DFE18E: mov edx, var_54
  loc_00DFE191: mov [eax+0000000Ch], edx
  loc_00DFE194: push 006DEFA8h ; "DropDownSymbol"
  loc_00DFE199: mov eax, var_84
  loc_00DFE19F: mov ecx, [eax]
  loc_00DFE1A1: mov edx, var_84
  loc_00DFE1A7: push edx
  loc_00DFE1A8: call [ecx+0000001Ch]
  loc_00DFE1AB: fnclex
  loc_00DFE1AD: mov var_70, eax
  loc_00DFE1B0: cmp var_70, 00000000h
  loc_00DFE1B4: jge 00DFE1D6h
  loc_00DFE1B6: push 0000001Ch
  loc_00DFE1B8: push 006DEB78h
  loc_00DFE1BD: mov eax, var_84
  loc_00DFE1C3: push eax
  loc_00DFE1C4: mov ecx, var_70
  loc_00DFE1C7: push ecx
  loc_00DFE1C8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE1CE: mov var_158, eax
  loc_00DFE1D4: jmp 00DFE1E0h
  loc_00DFE1D6: mov var_158, 00000000h
  loc_00DFE1E0: lea edx, var_40
  loc_00DFE1E3: push edx
  loc_00DFE1E4: call [004011D0h] ; __vbaI4Var
  loc_00DFE1EA: mov ecx, Me
  loc_00DFE1ED: mov [ecx+0000005Ch], eax
  loc_00DFE1F0: lea ecx, var_40
  loc_00DFE1F3: call [00401024h] ; __vbaFreeVar
  loc_00DFE1F9: mov var_4, 0000002Ah
  loc_00DFE200: mov var_58, 00000000h
  loc_00DFE207: mov var_60, 00000003h
  loc_00DFE20E: lea edx, var_40
  loc_00DFE211: push edx
  loc_00DFE212: mov eax, 00000010h
  loc_00DFE217: call 00402830h ; __vbaChkstk
  loc_00DFE21C: mov eax, esp
  loc_00DFE21E: mov ecx, var_60
  loc_00DFE221: mov [eax], ecx
  loc_00DFE223: mov edx, var_5C
  loc_00DFE226: mov [eax+00000004h], edx
  loc_00DFE229: mov ecx, var_58
  loc_00DFE22C: mov [eax+00000008h], ecx
  loc_00DFE22F: mov edx, var_54
  loc_00DFE232: mov [eax+0000000Ch], edx
  loc_00DFE235: push 006DEFE0h ; "ColorScheme"
  loc_00DFE23A: mov eax, var_84
  loc_00DFE240: mov ecx, [eax]
  loc_00DFE242: mov edx, var_84
  loc_00DFE248: push edx
  loc_00DFE249: call [ecx+0000001Ch]
  loc_00DFE24C: fnclex
  loc_00DFE24E: mov var_70, eax
  loc_00DFE251: cmp var_70, 00000000h
  loc_00DFE255: jge 00DFE277h
  loc_00DFE257: push 0000001Ch
  loc_00DFE259: push 006DEB78h
  loc_00DFE25E: mov eax, var_84
  loc_00DFE264: push eax
  loc_00DFE265: mov ecx, var_70
  loc_00DFE268: push ecx
  loc_00DFE269: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE26F: mov var_15C, eax
  loc_00DFE275: jmp 00DFE281h
  loc_00DFE277: mov var_15C, 00000000h
  loc_00DFE281: lea edx, var_40
  loc_00DFE284: push edx
  loc_00DFE285: call [004011D0h] ; __vbaI4Var
  loc_00DFE28B: mov ecx, Me
  loc_00DFE28E: mov [ecx+000000CCh], eax
  loc_00DFE294: lea ecx, var_40
  loc_00DFE297: call [00401024h] ; __vbaFreeVar
  loc_00DFE29D: mov var_4, 0000002Bh
  loc_00DFE2A4: mov edx, Me
  loc_00DFE2A7: mov ax, [edx+0000007Ch]
  loc_00DFE2AB: push eax
  loc_00DFE2AC: mov ecx, Me
  loc_00DFE2AF: mov edx, [ecx+00000010h]
  loc_00DFE2B2: mov eax, Me
  loc_00DFE2B5: mov ecx, [eax+00000010h]
  loc_00DFE2B8: mov eax, [ecx]
  loc_00DFE2BA: push edx
  loc_00DFE2BB: call [eax+00000094h]
  loc_00DFE2C1: fnclex
  loc_00DFE2C3: mov var_70, eax
  loc_00DFE2C6: cmp var_70, 00000000h
  loc_00DFE2CA: jge 00DFE2EFh
  loc_00DFE2CC: push 00000094h
  loc_00DFE2D1: push 006DCB00h
  loc_00DFE2D6: mov ecx, Me
  loc_00DFE2D9: mov edx, [ecx+00000010h]
  loc_00DFE2DC: push edx
  loc_00DFE2DD: mov eax, var_70
  loc_00DFE2E0: push eax
  loc_00DFE2E1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE2E7: mov var_160, eax
  loc_00DFE2ED: jmp 00DFE2F9h
  loc_00DFE2EF: mov var_160, 00000000h
  loc_00DFE2F9: mov var_4, 0000002Ch
  loc_00DFE300: mov ecx, Me
  loc_00DFE303: mov edx, [ecx]
  loc_00DFE305: mov eax, Me
  loc_00DFE308: push eax
  loc_00DFE309: call [edx+0000093Ch]
  loc_00DFE30F: mov var_4, 0000002Dh
  loc_00DFE316: lea ecx, var_68
  loc_00DFE319: push ecx
  loc_00DFE31A: mov edx, Me
  loc_00DFE31D: mov eax, [edx+00000010h]
  loc_00DFE320: mov ecx, Me
  loc_00DFE323: mov edx, [ecx+00000010h]
  loc_00DFE326: mov ecx, [edx]
  loc_00DFE328: push eax
  loc_00DFE329: call [ecx+00000108h]
  loc_00DFE32F: fnclex
  loc_00DFE331: mov var_70, eax
  loc_00DFE334: cmp var_70, 00000000h
  loc_00DFE338: jge 00DFE35Dh
  loc_00DFE33A: push 00000108h
  loc_00DFE33F: push 006DCB00h
  loc_00DFE344: mov edx, Me
  loc_00DFE347: mov eax, [edx+00000010h]
  loc_00DFE34A: push eax
  loc_00DFE34B: mov ecx, var_70
  loc_00DFE34E: push ecx
  loc_00DFE34F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE355: mov var_164, eax
  loc_00DFE35B: jmp 00DFE367h
  loc_00DFE35D: mov var_164, 00000000h
  loc_00DFE367: fld real4 ptr var_68
  loc_00DFE36A: call [00401208h] ; __vbaFpI4
  loc_00DFE370: mov edx, Me
  loc_00DFE373: mov [edx+000001D0h], eax
  loc_00DFE379: mov var_4, 0000002Eh
  loc_00DFE380: lea eax, var_68
  loc_00DFE383: push eax
  loc_00DFE384: mov ecx, Me
  loc_00DFE387: mov edx, [ecx+00000010h]
  loc_00DFE38A: mov eax, Me
  loc_00DFE38D: mov ecx, [eax+00000010h]
  loc_00DFE390: mov eax, [ecx]
  loc_00DFE392: push edx
  loc_00DFE393: call [eax+00000100h]
  loc_00DFE399: fnclex
  loc_00DFE39B: mov var_70, eax
  loc_00DFE39E: cmp var_70, 00000000h
  loc_00DFE3A2: jge 00DFE3C7h
  loc_00DFE3A4: push 00000100h
  loc_00DFE3A9: push 006DCB00h
  loc_00DFE3AE: mov ecx, Me
  loc_00DFE3B1: mov edx, [ecx+00000010h]
  loc_00DFE3B4: push edx
  loc_00DFE3B5: mov eax, var_70
  loc_00DFE3B8: push eax
  loc_00DFE3B9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE3BF: mov var_168, eax
  loc_00DFE3C5: jmp 00DFE3D1h
  loc_00DFE3C7: mov var_168, 00000000h
  loc_00DFE3D1: fld real4 ptr var_68
  loc_00DFE3D4: call [00401208h] ; __vbaFpI4
  loc_00DFE3DA: mov ecx, Me
  loc_00DFE3DD: mov [ecx+000001D4h], eax
  loc_00DFE3E3: mov var_4, 0000002Fh
  loc_00DFE3EA: lea edx, var_28
  loc_00DFE3ED: push edx
  loc_00DFE3EE: mov eax, Me
  loc_00DFE3F1: mov ecx, [eax+00000010h]
  loc_00DFE3F4: mov edx, Me
  loc_00DFE3F7: mov eax, [edx+00000010h]
  loc_00DFE3FA: mov edx, [eax]
  loc_00DFE3FC: push ecx
  loc_00DFE3FD: call [edx+00000198h]
  loc_00DFE403: fnclex
  loc_00DFE405: mov var_70, eax
  loc_00DFE408: cmp var_70, 00000000h
  loc_00DFE40C: jge 00DFE431h
  loc_00DFE40E: push 00000198h
  loc_00DFE413: push 006DCB00h
  loc_00DFE418: mov eax, Me
  loc_00DFE41B: mov ecx, [eax+00000010h]
  loc_00DFE41E: push ecx
  loc_00DFE41F: mov edx, var_70
  loc_00DFE422: push edx
  loc_00DFE423: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE429: mov var_16C, eax
  loc_00DFE42F: jmp 00DFE43Bh
  loc_00DFE431: mov var_16C, 00000000h
  loc_00DFE43B: push 00000000h
  loc_00DFE43D: push 006DEA44h ; "Hwnd"
  loc_00DFE442: mov eax, var_28
  loc_00DFE445: push eax
  loc_00DFE446: lea ecx, var_40
  loc_00DFE449: push ecx
  loc_00DFE44A: call [00401214h] ; __vbaLateMemCallLd
  loc_00DFE450: add esp, 00000010h
  loc_00DFE453: push eax
  loc_00DFE454: call [004011D0h] ; __vbaI4Var
  loc_00DFE45A: mov edx, Me
  loc_00DFE45D: mov [edx+00000074h], eax
  loc_00DFE460: lea ecx, var_28
  loc_00DFE463: call [00401254h] ; __vbaFreeObj
  loc_00DFE469: lea ecx, var_40
  loc_00DFE46C: call [00401024h] ; __vbaFreeVar
  loc_00DFE472: mov var_4, 00000030h
  loc_00DFE479: push 00000000h
  loc_00DFE47B: lea eax, var_84
  loc_00DFE481: push eax
  loc_00DFE482: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFE488: mov var_4, 00000031h
  loc_00DFE48F: mov ecx, Me
  loc_00DFE492: movsx edx, [ecx+00000052h]
  loc_00DFE496: test edx, edx
  loc_00DFE498: jz 00DFE4DEh
  loc_00DFE49A: mov var_4, 00000032h
  loc_00DFE4A1: push 00007F89h
  loc_00DFE4A6: push 00000000h
  loc_00DFE4A8: call 006DDD2Ch ; LoadCursor(%x1v, %x2v)
  loc_00DFE4AD: mov var_68, eax
  loc_00DFE4B0: call [00401074h] ; __vbaSetSystemError
  loc_00DFE4B6: mov eax, Me
  loc_00DFE4B9: mov ecx, var_68
  loc_00DFE4BC: mov [eax+00000054h], ecx
  loc_00DFE4BF: mov var_4, 00000033h
  loc_00DFE4C6: mov edx, Me
  loc_00DFE4C9: xor eax, eax
  loc_00DFE4CB: cmp [edx+00000054h], 00000000h
  loc_00DFE4CF: setz al
  loc_00DFE4D2: neg eax
  loc_00DFE4D4: not ax
  loc_00DFE4D7: mov ecx, Me
  loc_00DFE4DA: mov [ecx+00000052h], ax
  loc_00DFE4DE: mov var_4, 00000035h
  loc_00DFE4E5: mov edx, Me
  loc_00DFE4E8: mov eax, [edx]
  loc_00DFE4EA: mov ecx, Me
  loc_00DFE4ED: push ecx
  loc_00DFE4EE: call [eax+000009C4h]
  loc_00DFE4F4: mov var_70, eax
  loc_00DFE4F7: cmp var_70, 00000000h
  loc_00DFE4FB: jge 00DFE51Dh
  loc_00DFE4FD: push 000009C4h
  loc_00DFE502: push 006DD090h
  loc_00DFE507: mov edx, Me
  loc_00DFE50A: push edx
  loc_00DFE50B: mov eax, var_70
  loc_00DFE50E: push eax
  loc_00DFE50F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE515: mov var_170, eax
  loc_00DFE51B: jmp 00DFE527h
  loc_00DFE51D: mov var_170, 00000000h
  loc_00DFE527: mov var_4, 00000036h
  loc_00DFE52E: push 00000001h
  loc_00DFE530: call [004010A4h] ; __vbaOnError
  loc_00DFE536: mov var_4, 00000037h
  loc_00DFE53D: lea ecx, var_28
  loc_00DFE540: push ecx
  loc_00DFE541: mov edx, Me
  loc_00DFE544: mov eax, [edx]
  loc_00DFE546: mov ecx, Me
  loc_00DFE549: push ecx
  loc_00DFE54A: call [eax+000002B0h]
  loc_00DFE550: fnclex
  loc_00DFE552: mov var_70, eax
  loc_00DFE555: cmp var_70, 00000000h
  loc_00DFE559: jge 00DFE57Bh
  loc_00DFE55B: push 000002B0h
  loc_00DFE560: push 006DCB00h
  loc_00DFE565: mov edx, Me
  loc_00DFE568: push edx
  loc_00DFE569: mov eax, var_70
  loc_00DFE56C: push eax
  loc_00DFE56D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE573: mov var_174, eax
  loc_00DFE579: jmp 00DFE585h
  loc_00DFE57B: mov var_174, 00000000h
  loc_00DFE585: mov ecx, var_28
  loc_00DFE588: mov var_74, ecx
  loc_00DFE58B: lea edx, var_64
  loc_00DFE58E: push edx
  loc_00DFE58F: mov eax, var_74
  loc_00DFE592: mov ecx, [eax]
  loc_00DFE594: mov edx, var_74
  loc_00DFE597: push edx
  loc_00DFE598: call [ecx+0000003Ch]
  loc_00DFE59B: fnclex
  loc_00DFE59D: mov var_78, eax
  loc_00DFE5A0: cmp var_78, 00000000h
  loc_00DFE5A4: jge 00DFE5C3h
  loc_00DFE5A6: push 0000003Ch
  loc_00DFE5A8: push 006DEA70h
  loc_00DFE5AD: mov eax, var_74
  loc_00DFE5B0: push eax
  loc_00DFE5B1: mov ecx, var_78
  loc_00DFE5B4: push ecx
  loc_00DFE5B5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE5BB: mov var_178, eax
  loc_00DFE5C1: jmp 00DFE5CDh
  loc_00DFE5C3: mov var_178, 00000000h
  loc_00DFE5CD: mov dx, var_64
  loc_00DFE5D1: mov var_7C, dx
  loc_00DFE5D5: lea ecx, var_28
  loc_00DFE5D8: call [00401254h] ; __vbaFreeObj
  loc_00DFE5DE: movsx eax, var_7C
  loc_00DFE5E2: test eax, eax
  loc_00DFE5E4: jz 00DFE958h
  loc_00DFE5EA: mov var_4, 00000038h
  loc_00DFE5F1: lea ecx, var_64
  loc_00DFE5F4: push ecx
  loc_00DFE5F5: push 006DF020h ; "User32"
  loc_00DFE5FA: push 006DEFFCh ; "TrackMouseEvent"
  loc_00DFE5FF: mov edx, Me
  loc_00DFE602: mov eax, [edx]
  loc_00DFE604: mov ecx, Me
  loc_00DFE607: push ecx
  loc_00DFE608: call [eax+000009ECh]
  loc_00DFE60E: mov edx, Me
  loc_00DFE611: mov ax, var_64
  loc_00DFE615: mov [edx+00000040h], ax
  loc_00DFE619: mov var_4, 00000039h
  loc_00DFE620: mov ecx, Me
  loc_00DFE623: movsx edx, [ecx+00000040h]
  loc_00DFE627: test edx, edx
  loc_00DFE629: jnz 00DFE64Fh
  loc_00DFE62B: mov var_4, 0000003Ah
  loc_00DFE632: lea eax, var_64
  loc_00DFE635: push eax
  loc_00DFE636: push 006DF05Ch ; "ComCtl32"
  loc_00DFE63B: push 006DF034h ; "_TrackMouseEvent"
  loc_00DFE640: mov ecx, Me
  loc_00DFE643: mov edx, [ecx]
  loc_00DFE645: mov eax, Me
  loc_00DFE648: push eax
  loc_00DFE649: call [edx+000009ECh]
  loc_00DFE64F: mov var_4, 0000003Ch
  loc_00DFE656: mov ecx, Me
  loc_00DFE659: mov edx, [ecx+00000010h]
  loc_00DFE65C: push edx
  loc_00DFE65D: lea eax, var_88
  loc_00DFE663: push eax
  loc_00DFE664: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFE66A: mov var_4, 0000003Dh
  loc_00DFE671: lea ecx, var_68
  loc_00DFE674: push ecx
  loc_00DFE675: mov edx, var_88
  loc_00DFE67B: mov eax, [edx]
  loc_00DFE67D: mov ecx, var_88
  loc_00DFE683: push ecx
  loc_00DFE684: call [eax+00000058h]
  loc_00DFE687: fnclex
  loc_00DFE689: mov var_70, eax
  loc_00DFE68C: cmp var_70, 00000000h
  loc_00DFE690: jge 00DFE6B2h
  loc_00DFE692: push 00000058h
  loc_00DFE694: push 006DCB00h
  loc_00DFE699: mov edx, var_88
  loc_00DFE69F: push edx
  loc_00DFE6A0: mov eax, var_70
  loc_00DFE6A3: push eax
  loc_00DFE6A4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE6AA: mov var_17C, eax
  loc_00DFE6B0: jmp 00DFE6BCh
  loc_00DFE6B2: mov var_17C, 00000000h
  loc_00DFE6BC: lea ecx, var_6C
  loc_00DFE6BF: push ecx
  loc_00DFE6C0: mov edx, var_68
  loc_00DFE6C3: push edx
  loc_00DFE6C4: mov eax, Me
  loc_00DFE6C7: mov ecx, [eax]
  loc_00DFE6C9: mov edx, Me
  loc_00DFE6CC: push edx
  loc_00DFE6CD: call [ecx+00000A04h]
  loc_00DFE6D3: mov var_4, 0000003Eh
  loc_00DFE6DA: lea eax, var_68
  loc_00DFE6DD: push eax
  loc_00DFE6DE: mov ecx, Me
  loc_00DFE6E1: mov edx, [ecx+00000074h]
  loc_00DFE6E4: push edx
  loc_00DFE6E5: mov eax, Me
  loc_00DFE6E8: mov ecx, [eax]
  loc_00DFE6EA: mov edx, Me
  loc_00DFE6ED: push edx
  loc_00DFE6EE: call [ecx+00000A04h]
  loc_00DFE6F4: mov var_4, 0000003Fh
  loc_00DFE6FB: lea eax, var_68
  loc_00DFE6FE: push eax
  loc_00DFE6FF: mov ecx, var_88
  loc_00DFE705: mov edx, [ecx]
  loc_00DFE707: mov eax, var_88
  loc_00DFE70D: push eax
  loc_00DFE70E: call [edx+00000058h]
  loc_00DFE711: fnclex
  loc_00DFE713: mov var_70, eax
  loc_00DFE716: cmp var_70, 00000000h
  loc_00DFE71A: jge 00DFE73Ch
  loc_00DFE71C: push 00000058h
  loc_00DFE71E: push 006DCB00h
  loc_00DFE723: mov ecx, var_88
  loc_00DFE729: push ecx
  loc_00DFE72A: mov edx, var_70
  loc_00DFE72D: push edx
  loc_00DFE72E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE734: mov var_180, eax
  loc_00DFE73A: jmp 00DFE746h
  loc_00DFE73C: mov var_180, 00000000h
  loc_00DFE746: push 00000001h
  loc_00DFE748: push 000002A3h
  loc_00DFE74D: mov eax, var_68
  loc_00DFE750: push eax
  loc_00DFE751: mov ecx, Me
  loc_00DFE754: mov edx, [ecx]
  loc_00DFE756: mov eax, Me
  loc_00DFE759: push eax
  loc_00DFE75A: call [edx+00000A0Ch]
  loc_00DFE760: mov var_4, 00000040h
  loc_00DFE767: lea ecx, var_68
  loc_00DFE76A: push ecx
  loc_00DFE76B: mov edx, var_88
  loc_00DFE771: mov eax, [edx]
  loc_00DFE773: mov ecx, var_88
  loc_00DFE779: push ecx
  loc_00DFE77A: call [eax+00000058h]
  loc_00DFE77D: fnclex
  loc_00DFE77F: mov var_70, eax
  loc_00DFE782: cmp var_70, 00000000h
  loc_00DFE786: jge 00DFE7A8h
  loc_00DFE788: push 00000058h
  loc_00DFE78A: push 006DCB00h
  loc_00DFE78F: mov edx, var_88
  loc_00DFE795: push edx
  loc_00DFE796: mov eax, var_70
  loc_00DFE799: push eax
  loc_00DFE79A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE7A0: mov var_184, eax
  loc_00DFE7A6: jmp 00DFE7B2h
  loc_00DFE7A8: mov var_184, 00000000h
  loc_00DFE7B2: push 00000001h
  loc_00DFE7B4: push 0000031Ah
  loc_00DFE7B9: mov ecx, var_68
  loc_00DFE7BC: push ecx
  loc_00DFE7BD: mov edx, Me
  loc_00DFE7C0: mov eax, [edx]
  loc_00DFE7C2: mov ecx, Me
  loc_00DFE7C5: push ecx
  loc_00DFE7C6: call [eax+00000A0Ch]
  loc_00DFE7CC: mov var_4, 00000041h
  loc_00DFE7D3: lea edx, var_64
  loc_00DFE7D6: push edx
  loc_00DFE7D7: mov eax, Me
  loc_00DFE7DA: mov ecx, [eax]
  loc_00DFE7DC: mov edx, Me
  loc_00DFE7DF: push edx
  loc_00DFE7E0: call [ecx+000009E4h]
  loc_00DFE7E6: movsx eax, var_64
  loc_00DFE7EA: test eax, eax
  loc_00DFE7EC: jz 00DFE857h
  loc_00DFE7EE: mov var_4, 00000042h
  loc_00DFE7F5: lea ecx, var_68
  loc_00DFE7F8: push ecx
  loc_00DFE7F9: mov edx, var_88
  loc_00DFE7FF: mov eax, [edx]
  loc_00DFE801: mov ecx, var_88
  loc_00DFE807: push ecx
  loc_00DFE808: call [eax+00000058h]
  loc_00DFE80B: fnclex
  loc_00DFE80D: mov var_70, eax
  loc_00DFE810: cmp var_70, 00000000h
  loc_00DFE814: jge 00DFE836h
  loc_00DFE816: push 00000058h
  loc_00DFE818: push 006DCB00h
  loc_00DFE81D: mov edx, var_88
  loc_00DFE823: push edx
  loc_00DFE824: mov eax, var_70
  loc_00DFE827: push eax
  loc_00DFE828: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE82E: mov var_188, eax
  loc_00DFE834: jmp 00DFE840h
  loc_00DFE836: mov var_188, 00000000h
  loc_00DFE840: push 00000001h
  loc_00DFE842: push 00000015h
  loc_00DFE844: mov ecx, var_68
  loc_00DFE847: push ecx
  loc_00DFE848: mov edx, Me
  loc_00DFE84B: mov eax, [edx]
  loc_00DFE84D: mov ecx, Me
  loc_00DFE850: push ecx
  loc_00DFE851: call [eax+00000A0Ch]
  loc_00DFE857: mov var_4, 00000044h
  loc_00DFE85E: push FFFFFFFFh
  loc_00DFE860: call [004010A4h] ; __vbaOnError
  loc_00DFE866: mov var_4, 00000045h
  loc_00DFE86D: lea edx, var_28
  loc_00DFE870: push edx
  loc_00DFE871: mov eax, Me
  loc_00DFE874: mov ecx, [eax+00000010h]
  loc_00DFE877: mov edx, Me
  loc_00DFE87A: mov eax, [edx+00000010h]
  loc_00DFE87D: mov edx, [eax]
  loc_00DFE87F: push ecx
  loc_00DFE880: call [edx+00000198h]
  loc_00DFE886: fnclex
  loc_00DFE888: mov var_70, eax
  loc_00DFE88B: cmp var_70, 00000000h
  loc_00DFE88F: jge 00DFE8B4h
  loc_00DFE891: push 00000198h
  loc_00DFE896: push 006DCB00h
  loc_00DFE89B: mov eax, Me
  loc_00DFE89E: mov ecx, [eax+00000010h]
  loc_00DFE8A1: push ecx
  loc_00DFE8A2: mov edx, var_70
  loc_00DFE8A5: push edx
  loc_00DFE8A6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFE8AC: mov var_18C, eax
  loc_00DFE8B2: jmp 00DFE8BEh
  loc_00DFE8B4: mov var_18C, 00000000h
  loc_00DFE8BE: push 00000000h
  loc_00DFE8C0: push 006DF070h ; "MDIChild"
  loc_00DFE8C5: mov eax, var_28
  loc_00DFE8C8: push eax
  loc_00DFE8C9: lea ecx, var_40
  loc_00DFE8CC: push ecx
  loc_00DFE8CD: call [00401214h] ; __vbaLateMemCallLd
  loc_00DFE8D3: add esp, 00000010h
  loc_00DFE8D6: push eax
  loc_00DFE8D7: call [004010DCh] ; __vbaBoolVarNull
  loc_00DFE8DD: mov var_74, ax
  loc_00DFE8E1: lea ecx, var_28
  loc_00DFE8E4: call [00401254h] ; __vbaFreeObj
  loc_00DFE8EA: lea ecx, var_40
  loc_00DFE8ED: call [00401024h] ; __vbaFreeVar
  loc_00DFE8F3: movsx edx, var_74
  loc_00DFE8F7: test edx, edx
  loc_00DFE8F9: jz 00DFE921h
  loc_00DFE8FB: mov var_4, 00000046h
  loc_00DFE902: push 00000001h
  loc_00DFE904: push 00000086h
  loc_00DFE909: mov eax, Me
  loc_00DFE90C: mov ecx, [eax+00000074h]
  loc_00DFE90F: push ecx
  loc_00DFE910: mov edx, Me
  loc_00DFE913: mov eax, [edx]
  loc_00DFE915: mov ecx, Me
  loc_00DFE918: push ecx
  loc_00DFE919: call [eax+00000A0Ch]
  loc_00DFE91F: jmp 00DFE942h
  loc_00DFE921: mov var_4, 00000048h
  loc_00DFE928: push 00000001h
  loc_00DFE92A: push 00000006h
  loc_00DFE92C: mov edx, Me
  loc_00DFE92F: mov eax, [edx+00000074h]
  loc_00DFE932: push eax
  loc_00DFE933: mov ecx, Me
  loc_00DFE936: mov edx, [ecx]
  loc_00DFE938: mov eax, Me
  loc_00DFE93B: push eax
  loc_00DFE93C: call [edx+00000A0Ch]
  loc_00DFE942: mov var_4, 0000004Ah
  loc_00DFE949: push 00000000h
  loc_00DFE94B: lea ecx, var_88
  loc_00DFE951: push ecx
  loc_00DFE952: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFE958: call [00401098h] ; __vbaExitProc
  loc_00DFE95E: fwait
  loc_00DFE95F: push 00DFE9B4h
  loc_00DFE964: jmp 00DFE99Ah
  loc_00DFE966: lea ecx, var_24
  loc_00DFE969: call [00401258h] ; __vbaFreeStr
  loc_00DFE96F: lea edx, var_30
  loc_00DFE972: push edx
  loc_00DFE973: lea eax, var_2C
  loc_00DFE976: push eax
  loc_00DFE977: lea ecx, var_28
  loc_00DFE97A: push ecx
  loc_00DFE97B: push 00000003h
  loc_00DFE97D: call [00401048h] ; __vbaFreeObjList
  loc_00DFE983: add esp, 00000010h
  loc_00DFE986: lea edx, var_50
  loc_00DFE989: push edx
  loc_00DFE98A: lea eax, var_40
  loc_00DFE98D: push eax
  loc_00DFE98E: push 00000002h
  loc_00DFE990: call [00401038h] ; __vbaFreeVarList
  loc_00DFE996: add esp, 0000000Ch
  loc_00DFE999: ret
  loc_00DFE99A: lea ecx, var_88
  loc_00DFE9A0: push ecx
  loc_00DFE9A1: lea edx, var_84
  loc_00DFE9A7: push edx
  loc_00DFE9A8: push 00000002h
  loc_00DFE9AA: call [00401048h] ; __vbaFreeObjList
  loc_00DFE9B0: add esp, 0000000Ch
  loc_00DFE9B3: ret
  loc_00DFE9B4: mov eax, Me
  loc_00DFE9B7: mov ecx, [eax]
  loc_00DFE9B9: mov edx, Me
  loc_00DFE9BC: push edx
  loc_00DFE9BD: call [ecx+00000008h]
  loc_00DFE9C0: mov eax, var_10
  loc_00DFE9C3: mov ecx, var_20
  loc_00DFE9C6: mov fs:[00000000h], ecx
  loc_00DFE9CD: pop edi
  loc_00DFE9CE: pop esi
  loc_00DFE9CF: pop ebx
  loc_00DFE9D0: mov esp, ebp
  loc_00DFE9D2: pop ebp
  loc_00DFE9D3: retn 0008h
End Sub

Private Sub UserControl_UnknownEvent_21 'DFAAC0
  loc_00DFAAC0: push ebp
  loc_00DFAAC1: mov ebp, esp
  loc_00DFAAC3: sub esp, 0000000Ch
  loc_00DFAAC6: push 00402836h ; __vbaExceptHandler
  loc_00DFAACB: mov eax, fs:[00000000h]
  loc_00DFAAD1: push eax
  loc_00DFAAD2: mov fs:[00000000h], esp
  loc_00DFAAD9: sub esp, 00000028h
  loc_00DFAADC: push ebx
  loc_00DFAADD: push esi
  loc_00DFAADE: push edi
  loc_00DFAADF: mov var_C, esp
  loc_00DFAAE2: mov var_8, 004018D8h
  loc_00DFAAE9: mov esi, Me
  loc_00DFAAEC: mov eax, esi
  loc_00DFAAEE: and eax, 00000001h
  loc_00DFAAF1: mov var_4, eax
  loc_00DFAAF4: and esi, FFFFFFFEh
  loc_00DFAAF7: push esi
  loc_00DFAAF8: mov Me, esi
  loc_00DFAAFB: mov ecx, [esi]
  loc_00DFAAFD: call [ecx+00000004h]
  loc_00DFAB00: mov edx, [esi]
  loc_00DFAB02: or eax, FFFFFFFFh
  loc_00DFAB05: mov [esi+0000006Eh], ax
  loc_00DFAB09: mov [esi+0000007Ch], ax
  loc_00DFAB0D: lea eax, var_1C
  loc_00DFAB10: xor edi, edi
  loc_00DFAB12: push eax
  loc_00DFAB13: push esi
  loc_00DFAB14: mov var_18, edi
  loc_00DFAB17: mov var_1C, edi
  loc_00DFAB1A: mov var_20, edi
  loc_00DFAB1D: mov var_24, edi
  loc_00DFAB20: mov [esi+00000044h], 0000000Dh
  loc_00DFAB27: call [edx+000002B0h]
  loc_00DFAB2D: cmp eax, edi
  loc_00DFAB2F: fnclex
  loc_00DFAB31: jge 00DFAB49h
  loc_00DFAB33: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFAB39: push 000002B0h
  loc_00DFAB3E: push 006DCB00h
  loc_00DFAB43: push esi
  loc_00DFAB44: push eax
  loc_00DFAB45: call ebx
  loc_00DFAB47: jmp 00DFAB4Fh
  loc_00DFAB49: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFAB4F: mov eax, var_1C
  loc_00DFAB52: lea edx, var_18
  loc_00DFAB55: push edx
  loc_00DFAB56: push eax
  loc_00DFAB57: mov ecx, [eax]
  loc_00DFAB59: mov edi, eax
  loc_00DFAB5B: call [ecx+00000020h]
  loc_00DFAB5E: test eax, eax
  loc_00DFAB60: fnclex
  loc_00DFAB62: jge 00DFAB6Fh
  loc_00DFAB64: push 00000020h
  loc_00DFAB66: push 006DEA70h
  loc_00DFAB6B: push edi
  loc_00DFAB6C: push eax
  loc_00DFAB6D: call ebx
  loc_00DFAB6F: mov edx, var_18
  loc_00DFAB72: lea ecx, [esi+00000080h]
  loc_00DFAB78: call [004011B0h] ; __vbaStrCopy
  loc_00DFAB7E: lea ecx, var_18
  loc_00DFAB81: call [00401258h] ; __vbaFreeStr
  loc_00DFAB87: lea ecx, var_1C
  loc_00DFAB8A: call [00401254h] ; __vbaFreeObj
  loc_00DFAB90: mov eax, [esi+00000010h]
  loc_00DFAB93: push 006DEB18h ; "Tahoma"
  loc_00DFAB98: push eax
  loc_00DFAB99: mov ecx, [eax]
  loc_00DFAB9B: call [ecx+000000ACh]
  loc_00DFABA1: test eax, eax
  loc_00DFABA3: fnclex
  loc_00DFABA5: jge 00DFABB8h
  loc_00DFABA7: mov edx, [esi+00000010h]
  loc_00DFABAA: push 000000ACh
  loc_00DFABAF: push 006DCB00h
  loc_00DFABB4: push edx
  loc_00DFABB5: push eax
  loc_00DFABB6: call ebx
  loc_00DFABB8: mov eax, [esi+00000010h]
  loc_00DFABBB: lea edx, var_1C
  loc_00DFABBE: push edx
  loc_00DFABBF: push eax
  loc_00DFABC0: mov ecx, [eax]
  loc_00DFABC2: call [ecx+00000248h]
  loc_00DFABC8: test eax, eax
  loc_00DFABCA: fnclex
  loc_00DFABCC: jge 00DFABDFh
  loc_00DFABCE: mov ecx, [esi+00000010h]
  loc_00DFABD1: push 00000248h
  loc_00DFABD6: push 006DCB00h
  loc_00DFABDB: push ecx
  loc_00DFABDC: push eax
  loc_00DFABDD: call ebx
  loc_00DFABDF: mov eax, var_1C
  loc_00DFABE2: mov edi, [esi]
  loc_00DFABE4: lea edx, var_20
  loc_00DFABE7: push eax
  loc_00DFABE8: push edx
  loc_00DFABE9: mov var_1C, 00000000h
  loc_00DFABF0: call [004010ACh] ; __vbaObjSet
  loc_00DFABF6: push eax
  loc_00DFABF7: push esi
  loc_00DFABF8: call [edi+00000A34h]
  loc_00DFABFE: test eax, eax
  loc_00DFAC00: fnclex
  loc_00DFAC02: jge 00DFAC12h
  loc_00DFAC04: push 00000A34h
  loc_00DFAC09: push 006DD090h
  loc_00DFAC0E: push esi
  loc_00DFAC0F: push eax
  loc_00DFAC10: call ebx
  loc_00DFAC12: lea ecx, var_20
  loc_00DFAC15: call [00401254h] ; __vbaFreeObj
  loc_00DFAC1B: mov eax, [esi]
  loc_00DFAC1D: push 006DCC80h
  loc_00DFAC22: push esi
  loc_00DFAC23: call [eax+000009F4h]
  loc_00DFAC29: test eax, eax
  loc_00DFAC2B: jge 00DFAC3Bh
  loc_00DFAC2D: push 000009F4h
  loc_00DFAC32: push 006DD090h
  loc_00DFAC37: push esi
  loc_00DFAC38: push eax
  loc_00DFAC39: call ebx
  loc_00DFAC3B: mov edi, [00401144h] ; __vbaUI1I2
  loc_00DFAC41: mov ecx, 000000FFh
  loc_00DFAC46: call edi
  loc_00DFAC48: mov ecx, 000000FFh
  loc_00DFAC4D: mov [esi+00000158h], al
  loc_00DFAC53: call edi
  loc_00DFAC55: mov [esi+00000159h], al
  loc_00DFAC5B: mov eax, 00000001h
  loc_00DFAC60: mov [esi+00000164h], eax
  loc_00DFAC66: mov [esi+00000084h], eax
  loc_00DFAC6C: mov eax, [esi+00000010h]
  loc_00DFAC6F: lea edx, var_24
  loc_00DFAC72: mov [esi+0000009Ch], FFFFFFh
  loc_00DFAC7B: mov [esi+000000A0h], 00E0E0E0h
  loc_00DFAC85: mov ecx, [eax]
  loc_00DFAC87: push edx
  loc_00DFAC88: push eax
  loc_00DFAC89: call [ecx+00000108h]
  loc_00DFAC8F: test eax, eax
  loc_00DFAC91: fnclex
  loc_00DFAC93: jge 00DFACA6h
  loc_00DFAC95: mov ecx, [esi+00000010h]
  loc_00DFAC98: push 00000108h
  loc_00DFAC9D: push 006DCB00h
  loc_00DFACA2: push ecx
  loc_00DFACA3: push eax
  loc_00DFACA4: call ebx
  loc_00DFACA6: fld real4 ptr var_24
  loc_00DFACA9: mov edi, [00401208h] ; __vbaFpI4
  loc_00DFACAF: call edi
  loc_00DFACB1: mov [esi+000001D0h], eax
  loc_00DFACB7: mov eax, [esi+00000010h]
  loc_00DFACBA: lea ecx, var_24
  loc_00DFACBD: mov edx, [eax]
  loc_00DFACBF: push ecx
  loc_00DFACC0: push eax
  loc_00DFACC1: call [edx+00000100h]
  loc_00DFACC7: test eax, eax
  loc_00DFACC9: fnclex
  loc_00DFACCB: jge 00DFACDEh
  loc_00DFACCD: mov edx, [esi+00000010h]
  loc_00DFACD0: push 00000100h
  loc_00DFACD5: push 006DCB00h
  loc_00DFACDA: push edx
  loc_00DFACDB: push eax
  loc_00DFACDC: call ebx
  loc_00DFACDE: fld real4 ptr var_24
  loc_00DFACE1: call edi
  loc_00DFACE3: mov [esi+000001D4h], eax
  loc_00DFACE9: mov eax, [esi]
  loc_00DFACEB: push esi
  loc_00DFACEC: call [eax+000009B4h]
  loc_00DFACF2: mov ecx, [esi]
  loc_00DFACF4: push esi
  loc_00DFACF5: call [ecx+000009B8h]
  loc_00DFACFB: mov edx, [esi]
  loc_00DFACFD: push esi
  loc_00DFACFE: call [edx+00000338h]
  loc_00DFAD04: test eax, eax
  loc_00DFAD06: fnclex
  loc_00DFAD08: jge 00DFAD18h
  loc_00DFAD0A: push 00000338h
  loc_00DFAD0F: push 006DCB00h
  loc_00DFAD14: push esi
  loc_00DFAD15: push eax
  loc_00DFAD16: call ebx
  loc_00DFAD18: mov var_4, 00000000h
  loc_00DFAD1F: fwait
  loc_00DFAD20: push 00DFAD45h
  loc_00DFAD25: jmp 00DFAD44h
  loc_00DFAD27: lea ecx, var_18
  loc_00DFAD2A: call [00401258h] ; __vbaFreeStr
  loc_00DFAD30: lea eax, var_20
  loc_00DFAD33: lea ecx, var_1C
  loc_00DFAD36: push eax
  loc_00DFAD37: push ecx
  loc_00DFAD38: push 00000002h
  loc_00DFAD3A: call [00401048h] ; __vbaFreeObjList
  loc_00DFAD40: add esp, 0000000Ch
  loc_00DFAD43: ret
  loc_00DFAD44: ret
  loc_00DFAD45: mov eax, Me
  loc_00DFAD48: push eax
  loc_00DFAD49: mov edx, [eax]
  loc_00DFAD4B: call [edx+00000008h]
  loc_00DFAD4E: mov eax, var_4
  loc_00DFAD51: mov ecx, var_14
  loc_00DFAD54: pop edi
  loc_00DFAD55: pop esi
  loc_00DFAD56: mov fs:[00000000h], ecx
  loc_00DFAD5D: pop ebx
  loc_00DFAD5E: mov esp, ebp
  loc_00DFAD60: pop ebp
  loc_00DFAD61: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_25 'DFE9E0
  loc_00DFE9E0: push ebp
  loc_00DFE9E1: mov ebp, esp
  loc_00DFE9E3: sub esp, 0000000Ch
  loc_00DFE9E6: push 00402836h ; __vbaExceptHandler
  loc_00DFE9EB: mov eax, fs:[00000000h]
  loc_00DFE9F1: push eax
  loc_00DFE9F2: mov fs:[00000000h], esp
  loc_00DFE9F9: sub esp, 00000030h
  loc_00DFE9FC: push ebx
  loc_00DFE9FD: push esi
  loc_00DFE9FE: push edi
  loc_00DFE9FF: mov var_C, esp
  loc_00DFEA02: mov var_8, 00401AB8h
  loc_00DFEA09: mov esi, Me
  loc_00DFEA0C: mov eax, esi
  loc_00DFEA0E: and eax, 00000001h
  loc_00DFEA11: mov var_4, eax
  loc_00DFEA14: and esi, FFFFFFFEh
  loc_00DFEA17: push esi
  loc_00DFEA18: mov Me, esi
  loc_00DFEA1B: mov ecx, [esi]
  loc_00DFEA1D: call [ecx+00000004h]
  loc_00DFEA20: mov eax, [esi+00000010h]
  loc_00DFEA23: xor edi, edi
  loc_00DFEA25: lea ecx, var_18
  loc_00DFEA28: mov var_18, edi
  loc_00DFEA2B: mov edx, [eax]
  loc_00DFEA2D: push ecx
  loc_00DFEA2E: push eax
  loc_00DFEA2F: mov ebx, 00000008h
  loc_00DFEA34: call [edx+000002C0h]
  loc_00DFEA3A: test eax, eax
  loc_00DFEA3C: fnclex
  loc_00DFEA3E: jge 00DFEA55h
  loc_00DFEA40: mov edx, [esi+00000010h]
  loc_00DFEA43: push 000002C0h
  loc_00DFEA48: push 006DCB00h
  loc_00DFEA4D: push edx
  loc_00DFEA4E: push eax
  loc_00DFEA4F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFEA55: mov ecx, var_24
  loc_00DFEA58: sub esp, 00000010h
  loc_00DFEA5B: mov eax, esp
  loc_00DFEA5D: mov edx, var_1C
  loc_00DFEA60: push 006DEAE0h ; "ToolTipText"
  loc_00DFEA65: mov [eax], ebx
  loc_00DFEA67: mov [eax+00000004h], ecx
  loc_00DFEA6A: mov [eax+00000008h], edi
  loc_00DFEA6D: mov [eax+0000000Ch], edx
  loc_00DFEA70: mov eax, var_18
  loc_00DFEA73: push eax
  loc_00DFEA74: call [00401094h] ; __vbaLateMemSt
  loc_00DFEA7A: lea ecx, var_18
  loc_00DFEA7D: call [00401254h] ; __vbaFreeObj
  loc_00DFEA83: mov var_4, 00000000h
  loc_00DFEA8A: push 00DFEA9Ch
  loc_00DFEA8F: jmp 00DFEA9Bh
  loc_00DFEA91: lea ecx, var_18
  loc_00DFEA94: call [00401254h] ; __vbaFreeObj
  loc_00DFEA9A: ret
  loc_00DFEA9B: ret
  loc_00DFEA9C: mov eax, Me
  loc_00DFEA9F: push eax
  loc_00DFEAA0: mov ecx, [eax]
  loc_00DFEAA2: call [ecx+00000008h]
  loc_00DFEAA5: mov eax, var_4
  loc_00DFEAA8: mov ecx, var_14
  loc_00DFEAAB: pop edi
  loc_00DFEAAC: pop esi
  loc_00DFEAAD: mov fs:[00000000h], ecx
  loc_00DFEAB4: pop ebx
  loc_00DFEAB5: mov esp, ebp
  loc_00DFEAB7: pop ebp
  loc_00DFEAB8: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_26 'DFA7F0
  loc_00DFA7F0: push ebp
  loc_00DFA7F1: mov ebp, esp
  loc_00DFA7F3: sub esp, 0000000Ch
  loc_00DFA7F6: push 00402836h ; __vbaExceptHandler
  loc_00DFA7FB: mov eax, fs:[00000000h]
  loc_00DFA801: push eax
  loc_00DFA802: mov fs:[00000000h], esp
  loc_00DFA809: sub esp, 00000030h
  loc_00DFA80C: push ebx
  loc_00DFA80D: push esi
  loc_00DFA80E: push edi
  loc_00DFA80F: mov var_C, esp
  loc_00DFA812: mov var_8, 004018B8h
  loc_00DFA819: mov esi, Me
  loc_00DFA81C: mov eax, esi
  loc_00DFA81E: and eax, 00000001h
  loc_00DFA821: mov var_4, eax
  loc_00DFA824: and esi, FFFFFFFEh
  loc_00DFA827: push esi
  loc_00DFA828: mov Me, esi
  loc_00DFA82B: mov ecx, [esi]
  loc_00DFA82D: call [ecx+00000004h]
  loc_00DFA830: mov eax, [esi+00000010h]
  loc_00DFA833: mov edi, [esi+000000F8h]
  loc_00DFA839: lea ecx, var_18
  loc_00DFA83C: mov var_18, 00000000h
  loc_00DFA843: mov edx, [eax]
  loc_00DFA845: push ecx
  loc_00DFA846: push eax
  loc_00DFA847: mov ebx, 00000008h
  loc_00DFA84C: call [edx+000002C0h]
  loc_00DFA852: test eax, eax
  loc_00DFA854: fnclex
  loc_00DFA856: jge 00DFA86Dh
  loc_00DFA858: mov edx, [esi+00000010h]
  loc_00DFA85B: push 000002C0h
  loc_00DFA860: push 006DCB00h
  loc_00DFA865: push edx
  loc_00DFA866: push eax
  loc_00DFA867: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFA86D: mov ecx, var_24
  loc_00DFA870: sub esp, 00000010h
  loc_00DFA873: mov eax, esp
  loc_00DFA875: mov edx, var_1C
  loc_00DFA878: push 006DEAE0h ; "ToolTipText"
  loc_00DFA87D: mov [eax], ebx
  loc_00DFA87F: mov [eax+00000004h], ecx
  loc_00DFA882: mov [eax+00000008h], edi
  loc_00DFA885: mov [eax+0000000Ch], edx
  loc_00DFA888: mov eax, var_18
  loc_00DFA88B: push eax
  loc_00DFA88C: call [00401094h] ; __vbaLateMemSt
  loc_00DFA892: lea ecx, var_18
  loc_00DFA895: call [00401254h] ; __vbaFreeObj
  loc_00DFA89B: mov var_4, 00000000h
  loc_00DFA8A2: push 00DFA8B4h
  loc_00DFA8A7: jmp 00DFA8B3h
  loc_00DFA8A9: lea ecx, var_18
  loc_00DFA8AC: call [00401254h] ; __vbaFreeObj
  loc_00DFA8B2: ret
  loc_00DFA8B3: ret
  loc_00DFA8B4: mov eax, Me
  loc_00DFA8B7: push eax
  loc_00DFA8B8: mov ecx, [eax]
  loc_00DFA8BA: call [ecx+00000008h]
  loc_00DFA8BD: mov eax, var_4
  loc_00DFA8C0: mov ecx, var_14
  loc_00DFA8C3: pop edi
  loc_00DFA8C4: pop esi
  loc_00DFA8C5: mov fs:[00000000h], ecx
  loc_00DFA8CC: pop ebx
  loc_00DFA8CD: mov esp, ebp
  loc_00DFA8CF: pop ebp
  loc_00DFA8D0: retn 0004h
End Sub

Private Sub UserControl_UnknownEvent_27 'DFA520
  loc_00DFA520: push ebp
  loc_00DFA521: mov ebp, esp
  loc_00DFA523: sub esp, 0000000Ch
  loc_00DFA526: push 00402836h ; __vbaExceptHandler
  loc_00DFA52B: mov eax, fs:[00000000h]
  loc_00DFA531: push eax
  loc_00DFA532: mov fs:[00000000h], esp
  loc_00DFA539: sub esp, 0000001Ch
  loc_00DFA53C: push ebx
  loc_00DFA53D: push esi
  loc_00DFA53E: push edi
  loc_00DFA53F: mov var_C, esp
  loc_00DFA542: mov var_8, 00401898h
  loc_00DFA549: mov esi, Me
  loc_00DFA54C: mov eax, esi
  loc_00DFA54E: and eax, 00000001h
  loc_00DFA551: mov var_4, eax
  loc_00DFA554: and esi, FFFFFFFEh
  loc_00DFA557: push esi
  loc_00DFA558: mov Me, esi
  loc_00DFA55B: mov ecx, [esi]
  loc_00DFA55D: call [ecx+00000004h]
  loc_00DFA560: mov edx, [esi]
  loc_00DFA562: lea eax, var_18
  loc_00DFA565: xor ebx, ebx
  loc_00DFA567: push eax
  loc_00DFA568: push esi
  loc_00DFA569: mov var_18, ebx
  loc_00DFA56C: mov var_1C, ebx
  loc_00DFA56F: call [edx+000002B0h]
  loc_00DFA575: cmp eax, ebx
  loc_00DFA577: fnclex
  loc_00DFA579: jge 00DFA58Dh
  loc_00DFA57B: push 000002B0h
  loc_00DFA580: push 006DCB00h
  loc_00DFA585: push esi
  loc_00DFA586: push eax
  loc_00DFA587: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFA58D: mov eax, var_18
  loc_00DFA590: lea edx, var_1C
  loc_00DFA593: push edx
  loc_00DFA594: push eax
  loc_00DFA595: mov ecx, [eax]
  loc_00DFA597: mov edi, eax
  loc_00DFA599: call [ecx+0000004Ch]
  loc_00DFA59C: cmp eax, ebx
  loc_00DFA59E: fnclex
  loc_00DFA5A0: jge 00DFA5B1h
  loc_00DFA5A2: push 0000004Ch
  loc_00DFA5A4: push 006DEA70h
  loc_00DFA5A9: push edi
  loc_00DFA5AA: push eax
  loc_00DFA5AB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFA5B1: mov ax, var_1C
  loc_00DFA5B5: lea ecx, var_18
  loc_00DFA5B8: mov [esi+00000058h], ax
  loc_00DFA5BC: call [00401254h] ; __vbaFreeObj
  loc_00DFA5C2: mov edi, Cancel
  loc_00DFA5C5: mov ebx, [00401104h] ; __vbaStrCmp
  loc_00DFA5CB: mov ecx, [edi]
  loc_00DFA5CD: push ecx
  loc_00DFA5CE: push 006DEAA4h ; "DisplayAsDefault"
  loc_00DFA5D3: call ebx
  loc_00DFA5D5: test eax, eax
  loc_00DFA5D7: jnz 00DFA5E2h
  loc_00DFA5D9: mov edx, [esi]
  loc_00DFA5DB: push esi
  loc_00DFA5DC: call [edx+00000910h]
  loc_00DFA5E2: mov eax, [edi]
  loc_00DFA5E4: push eax
  loc_00DFA5E5: push 006DEACCh ; "BackColor"
  loc_00DFA5EA: call ebx
  loc_00DFA5EC: test eax, eax
  loc_00DFA5EE: jnz 00DFA5F9h
  loc_00DFA5F0: mov ecx, [esi]
  loc_00DFA5F2: push esi
  loc_00DFA5F3: call [ecx+00000910h]
  loc_00DFA5F9: mov var_4, 00000000h
  loc_00DFA600: push 00DFA612h
  loc_00DFA605: jmp 00DFA611h
  loc_00DFA607: lea ecx, var_18
  loc_00DFA60A: call [00401254h] ; __vbaFreeObj
  loc_00DFA610: ret
  loc_00DFA611: ret
  loc_00DFA612: mov eax, Me
  loc_00DFA615: push eax
  loc_00DFA616: mov edx, [eax]
  loc_00DFA618: call [edx+00000008h]
  loc_00DFA61B: mov eax, var_4
  loc_00DFA61E: mov ecx, var_14
  loc_00DFA621: pop edi
  loc_00DFA622: pop esi
  loc_00DFA623: mov fs:[00000000h], ecx
  loc_00DFA62A: pop ebx
  loc_00DFA62B: mov esp, ebp
  loc_00DFA62D: pop ebp
  loc_00DFA62E: retn 0008h
End Sub

Private Sub UserControl_UnknownEvent_28 'DFA410
  loc_00DFA410: push ebp
  loc_00DFA411: mov ebp, esp
  loc_00DFA413: sub esp, 0000000Ch
  loc_00DFA416: push 00402836h ; __vbaExceptHandler
  loc_00DFA41B: mov eax, fs:[00000000h]
  loc_00DFA421: push eax
  loc_00DFA422: mov fs:[00000000h], esp
  loc_00DFA429: sub esp, 00000008h
  loc_00DFA42C: push ebx
  loc_00DFA42D: push esi
  loc_00DFA42E: push edi
  loc_00DFA42F: mov var_C, esp
  loc_00DFA432: mov var_8, 00401890h
  loc_00DFA439: mov esi, Me
  loc_00DFA43C: mov eax, esi
  loc_00DFA43E: and eax, 00000001h
  loc_00DFA441: mov var_4, eax
  loc_00DFA444: and esi, FFFFFFFEh
  loc_00DFA447: push esi
  loc_00DFA448: mov Me, esi
  loc_00DFA44B: mov ecx, [esi]
  loc_00DFA44D: call [ecx+00000004h]
  loc_00DFA450: xor edi, edi
  loc_00DFA452: cmp [esi+0000007Ch], di
  loc_00DFA456: jz 00DFA4F8h
  loc_00DFA45C: cmp [esi+000000A8h], di
  loc_00DFA463: jz 00DFA470h
  loc_00DFA465: mov [esi+000000A8h], di
  loc_00DFA46C: mov [esi+0000004Ch], di
  loc_00DFA470: mov eax, [esi+00000064h]
  loc_00DFA473: cmp eax, 00000001h
  loc_00DFA476: jnz 00DFA4A9h
  loc_00DFA478: mov edx, Cancel
  loc_00DFA47B: xor ecx, ecx
  loc_00DFA47D: mov ax, [edx]
  loc_00DFA480: cmp ax, 001Bh
  loc_00DFA484: setnz cl
  loc_00DFA487: xor edx, edx
  loc_00DFA489: cmp ax, 000Dh
  loc_00DFA48D: setnz dl
  loc_00DFA490: test edx, ecx
  loc_00DFA492: jz 00DFA4F8h
  loc_00DFA494: mov ax, [esi+0000006Ch]
  loc_00DFA498: not ax
  loc_00DFA49B: cmp ax, di
  loc_00DFA49E: mov [esi+0000006Ch], ax
  loc_00DFA4A2: jnz 00DFA4D9h
  loc_00DFA4A4: mov [esi+00000048h], edi
  loc_00DFA4A7: jmp 00DFA4D9h
  loc_00DFA4A9: cmp eax, 00000002h
  loc_00DFA4AC: jnz 00DFA4E2h
  loc_00DFA4AE: mov ecx, Cancel
  loc_00DFA4B1: xor edx, edx
  loc_00DFA4B3: mov ax, [ecx]
  loc_00DFA4B6: cmp ax, 001Bh
  loc_00DFA4BA: setnz dl
  loc_00DFA4BD: xor ecx, ecx
  loc_00DFA4BF: cmp ax, 000Dh
  loc_00DFA4C3: setnz cl
  loc_00DFA4C6: test ecx, edx
  loc_00DFA4C8: jz 00DFA4F8h
  loc_00DFA4CA: mov edx, [esi]
  loc_00DFA4CC: push esi
  loc_00DFA4CD: call [edx+00000938h]
  loc_00DFA4D3: mov [esi+0000006Ch], FFFFFFh
  loc_00DFA4D9: mov eax, [esi]
  loc_00DFA4DB: push esi
  loc_00DFA4DC: call [eax+00000910h]
  loc_00DFA4E2: call [004010BCh] ; rtcDoEvents
  loc_00DFA4E8: push edi
  loc_00DFA4E9: push FFFFFDA8h
  loc_00DFA4EE: push esi
  loc_00DFA4EF: call [00401044h] ; __vbaRaiseEvent
  loc_00DFA4F5: add esp, 0000000Ch
  loc_00DFA4F8: mov var_4, edi
  loc_00DFA4FB: mov eax, Me
  loc_00DFA4FE: push eax
  loc_00DFA4FF: mov ecx, [eax]
  loc_00DFA501: call [ecx+00000008h]
  loc_00DFA504: mov eax, var_4
  loc_00DFA507: mov ecx, var_14
  loc_00DFA50A: pop edi
  loc_00DFA50B: pop esi
  loc_00DFA50C: mov fs:[00000000h], ecx
  loc_00DFA513: pop ebx
  loc_00DFA514: mov esp, ebp
  loc_00DFA516: pop ebp
  loc_00DFA517: retn 0008h
End Sub

Public Sub Subclass_WndProc(bBefore, bHandled, lReturn, lHwnd, uMsg, wParam, lParam) 'E00590
  loc_00E00590: push ebp
  loc_00E00591: mov ebp, esp
  loc_00E00593: sub esp, 0000000Ch
  loc_00E00596: push 00402836h ; __vbaExceptHandler
  loc_00E0059B: mov eax, fs:[00000000h]
  loc_00E005A1: push eax
  loc_00E005A2: mov fs:[00000000h], esp
  loc_00E005A9: sub esp, 0000000Ch
  loc_00E005AC: push ebx
  loc_00E005AD: push esi
  loc_00E005AE: push edi
  loc_00E005AF: mov var_C, esp
  loc_00E005B2: mov var_8, 00401BA0h
  loc_00E005B9: xor edi, edi
  loc_00E005BB: mov var_4, edi
  loc_00E005BE: mov esi, Me
  loc_00E005C1: push esi
  loc_00E005C2: mov eax, [esi]
  loc_00E005C4: call [eax+00000004h]
  loc_00E005C7: mov ecx, uMsg
  loc_00E005CA: mov eax, [ecx]
  loc_00E005CC: cmp eax, 00000086h
  loc_00E005D1: jg 00E0064Dh
  loc_00E005D3: jz 00E005F1h
  loc_00E005D5: cmp eax, 00000006h
  loc_00E005D8: jz 00E005F1h
  loc_00E005DA: cmp eax, 00000015h
  loc_00E005DD: jnz 00E006BAh
  loc_00E005E3: mov edx, [esi]
  loc_00E005E5: push esi
  loc_00E005E6: call [edx+00000910h]
  loc_00E005EC: jmp 00E006BAh
  loc_00E005F1: mov eax, wParam
  loc_00E005F4: cmp [eax], edi
  loc_00E005F6: jz 00E00625h
  loc_00E005F8: mov eax, [esi+00000048h]
  loc_00E005FB: mov [esi+00000070h], FFFFFFh
  loc_00E00601: cmp eax, edi
  loc_00E00603: jz 00E00608h
  loc_00E00605: mov [esi+00000048h], edi
  loc_00E00608: cmp [esi+00000058h], di
  loc_00E0060C: jz 00E00617h
  loc_00E0060E: mov ecx, [esi]
  loc_00E00610: push esi
  loc_00E00611: call [ecx+00000910h]
  loc_00E00617: mov edx, [esi]
  loc_00E00619: push esi
  loc_00E0061A: call [edx+00000910h]
  loc_00E00620: jmp 00E006BAh
  loc_00E00625: mov eax, [esi+00000048h]
  loc_00E00628: mov [esi+0000004Ch], di
  loc_00E0062C: cmp eax, edi
  loc_00E0062E: mov [esi+000000A8h], di
  loc_00E00635: mov [esi+00000050h], di
  loc_00E00639: mov [esi+00000070h], di
  loc_00E0063D: jz 00E00642h
  loc_00E0063F: mov [esi+00000048h], edi
  loc_00E00642: mov eax, [esi]
  loc_00E00644: push esi
  loc_00E00645: call [eax+00000910h]
  loc_00E0064B: jmp 00E006BAh
  loc_00E0064D: sub eax, 000002A3h
  loc_00E00652: jz 00E00664h
  loc_00E00654: sub eax, 00000077h
  loc_00E00657: jnz 00E006BAh
  loc_00E00659: mov ecx, [esi]
  loc_00E0065B: push esi
  loc_00E0065C: call [ecx+00000910h]
  loc_00E00662: jmp 00E006BAh
  loc_00E00664: cmp [esi+000000E0h], di
  loc_00E0066B: mov [esi+0000004Eh], di
  loc_00E0066F: jz 00E00693h
  loc_00E00671: cmp [esi+000000E4h], di
  loc_00E00678: jz 00E0068Ch
  loc_00E0067A: mov [esi+000000E4h], di
  loc_00E00681: mov [esi+000000E2h], FFFFFFh
  loc_00E0068A: jmp 00E006BAh
  loc_00E0068C: mov [esi+000000E2h], di
  loc_00E00693: cmp [esi+000000A8h], di
  loc_00E0069A: jnz 00E006BAh
  loc_00E0069C: cmp [esi+00000048h], edi
  loc_00E0069F: jz 00E006ADh
  loc_00E006A1: mov edx, [esi]
  loc_00E006A3: push esi
  loc_00E006A4: mov [esi+00000048h], edi
  loc_00E006A7: call [edx+00000910h]
  loc_00E006AD: push edi
  loc_00E006AE: push 00000002h
  loc_00E006B0: push esi
  loc_00E006B1: call [00401044h] ; __vbaRaiseEvent
  loc_00E006B7: add esp, 0000000Ch
  loc_00E006BA: mov eax, Me
  loc_00E006BD: push eax
  loc_00E006BE: mov ecx, [eax]
  loc_00E006C0: call [ecx+00000008h]
  loc_00E006C3: mov eax, var_4
  loc_00E006C6: mov ecx, var_14
  loc_00E006C9: pop edi
  loc_00E006CA: pop esi
  loc_00E006CB: mov fs:[00000000h], ecx
  loc_00E006D2: pop ebx
  loc_00E006D3: mov esp, ebp
  loc_00E006D5: pop ebp
  loc_00E006D6: retn 0020h
End Sub

Public Sub SetPopupMenu(menu, Align, Flags, DefaultMenu) 'E006E0
  loc_00E006E0: push ebp
  loc_00E006E1: mov ebp, esp
  loc_00E006E3: sub esp, 0000000Ch
  loc_00E006E6: push 00402836h ; __vbaExceptHandler
  loc_00E006EB: mov eax, fs:[00000000h]
  loc_00E006F1: push eax
  loc_00E006F2: mov fs:[00000000h], esp
  loc_00E006F9: sub esp, 0000001Ch
  loc_00E006FC: push ebx
  loc_00E006FD: push esi
  loc_00E006FE: push edi
  loc_00E006FF: mov var_C, esp
  loc_00E00702: mov var_8, 00401BA8h
  loc_00E00709: xor edi, edi
  loc_00E0070B: mov var_4, edi
  loc_00E0070E: mov esi, Me
  loc_00E00711: push esi
  loc_00E00712: mov eax, [esi]
  loc_00E00714: call [eax+00000004h]
  loc_00E00717: mov var_18, edi
  loc_00E0071A: mov var_28, edi
  loc_00E0071D: push edi
  loc_00E0071E: mov edi, menu
  loc_00E00721: mov ecx, [edi]
  loc_00E00723: push ecx
  loc_00E00724: call [0040114Ch] ; __vbaObjIs
  loc_00E0072A: test ax, ax
  loc_00E0072D: jnz 00E007B7h
  loc_00E00733: mov edx, [edi]
  loc_00E00735: push 006DF0D4h
  loc_00E0073A: push edx
  loc_00E0073B: call [00401188h] ; __vbaCheckType
  loc_00E00741: test ax, ax
  loc_00E00744: jz 00E007B7h
  loc_00E00746: mov eax, [edi]
  loc_00E00748: push 006DF0D4h
  loc_00E0074D: push eax
  loc_00E0074E: call [00401224h] ; __vbaCastObj
  loc_00E00754: lea ecx, var_18
  loc_00E00757: push eax
  loc_00E00758: push ecx
  loc_00E00759: call [004010ACh] ; __vbaObjSet
  loc_00E0075F: lea edx, [esi+000000E8h]
  loc_00E00765: push eax
  loc_00E00766: push edx
  loc_00E00767: call [004010B4h] ; __vbaObjSetAddref
  loc_00E0076D: lea ecx, var_18
  loc_00E00770: call [00401254h] ; __vbaFreeObj
  loc_00E00776: mov eax, Align
  loc_00E00779: mov edx, Flags
  loc_00E0077C: mov edi, [00401020h] ; __vbaVarVargNofree
  loc_00E00782: mov ecx, [eax]
  loc_00E00784: mov [esi+000000ECh], ecx
  loc_00E0078A: lea ecx, var_28
  loc_00E0078D: call edi
  loc_00E0078F: push eax
  loc_00E00790: call [004011D0h] ; __vbaI4Var
  loc_00E00796: mov ebx, DefaultMenu
  loc_00E00799: lea ecx, var_28
  loc_00E0079C: mov edx, ebx
  loc_00E0079E: mov [esi+000000F0h], eax
  loc_00E007A4: call edi
  loc_00E007A6: push eax
  loc_00E007A7: push ebx
  loc_00E007A8: call [0040105Ch] ; __vbaVarSetVarAddref
  loc_00E007AE: mov [esi+000000E0h], FFFFFFh
  loc_00E007B7: push 00E007C9h
  loc_00E007BC: jmp 00E007C8h
  loc_00E007BE: lea ecx, var_18
  loc_00E007C1: call [00401254h] ; __vbaFreeObj
  loc_00E007C7: ret
  loc_00E007C8: ret
  loc_00E007C9: mov eax, Me
  loc_00E007CC: push eax
  loc_00E007CD: mov edx, [eax]
  loc_00E007CF: call [edx+00000008h]
  loc_00E007D2: mov eax, var_4
  loc_00E007D5: mov ecx, var_14
  loc_00E007D8: pop edi
  loc_00E007D9: pop esi
  loc_00E007DA: mov fs:[00000000h], ecx
  loc_00E007E1: pop ebx
  loc_00E007E2: mov esp, ebp
  loc_00E007E4: pop ebp
  loc_00E007E5: retn 0014h
End Sub

Public Sub UnsetPopupMenu() 'E007F0
  loc_00E007F0: push ebp
  loc_00E007F1: mov ebp, esp
  loc_00E007F3: sub esp, 0000000Ch
  loc_00E007F6: push 00402836h ; __vbaExceptHandler
  loc_00E007FB: mov eax, fs:[00000000h]
  loc_00E00801: push eax
  loc_00E00802: mov fs:[00000000h], esp
  loc_00E00809: sub esp, 0000000Ch
  loc_00E0080C: push ebx
  loc_00E0080D: push esi
  loc_00E0080E: push edi
  loc_00E0080F: mov var_C, esp
  loc_00E00812: mov var_8, 00401BB8h
  loc_00E00819: xor edi, edi
  loc_00E0081B: mov var_4, edi
  loc_00E0081E: mov esi, Me
  loc_00E00821: push esi
  loc_00E00822: mov eax, [esi]
  loc_00E00824: call [eax+00000004h]
  loc_00E00827: push 006DF0D4h
  loc_00E0082C: mov var_18, edi
  loc_00E0082F: push edi
  loc_00E00830: mov edi, [00401224h] ; __vbaCastObj
  loc_00E00836: call edi
  loc_00E00838: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0083E: lea ecx, var_18
  loc_00E00841: push eax
  loc_00E00842: push ecx
  loc_00E00843: call ebx
  loc_00E00845: lea edx, [esi+000000E8h]
  loc_00E0084B: push eax
  loc_00E0084C: push edx
  loc_00E0084D: call [004010B4h] ; __vbaObjSetAddref
  loc_00E00853: lea ecx, var_18
  loc_00E00856: call [00401254h] ; __vbaFreeObj
  loc_00E0085C: push 006DF0D4h
  loc_00E00861: push 00000000h
  loc_00E00863: call edi
  loc_00E00865: push eax
  loc_00E00866: lea eax, var_18
  loc_00E00869: push eax
  loc_00E0086A: call ebx
  loc_00E0086C: lea ecx, [esi+000000F4h]
  loc_00E00872: push eax
  loc_00E00873: push ecx
  loc_00E00874: call [004010B4h] ; __vbaObjSetAddref
  loc_00E0087A: lea ecx, var_18
  loc_00E0087D: call [00401254h] ; __vbaFreeObj
  loc_00E00883: xor eax, eax
  loc_00E00885: mov [esi+000000E0h], ax
  loc_00E0088C: mov [esi+000000E2h], ax
  loc_00E00893: push 00E008A5h
  loc_00E00898: jmp 00E008A4h
  loc_00E0089A: lea ecx, var_18
  loc_00E0089D: call [00401254h] ; __vbaFreeObj
  loc_00E008A3: ret
  loc_00E008A4: ret
  loc_00E008A5: mov eax, Me
  loc_00E008A8: push eax
  loc_00E008A9: mov edx, [eax]
  loc_00E008AB: call [edx+00000008h]
  loc_00E008AE: mov eax, var_4
  loc_00E008B1: mov ecx, var_14
  loc_00E008B4: pop edi
  loc_00E008B5: pop esi
  loc_00E008B6: mov fs:[00000000h], ecx
  loc_00E008BD: pop ebx
  loc_00E008BE: mov esp, ebp
  loc_00E008C0: pop ebp
  loc_00E008C1: retn 0004h
End Sub

Public Sub OpenWebsite(sAddress) 'E008D0
  loc_00E008D0: push ebp
  loc_00E008D1: mov ebp, esp
  loc_00E008D3: sub esp, 0000000Ch
  loc_00E008D6: push 00402836h ; __vbaExceptHandler
  loc_00E008DB: mov eax, fs:[00000000h]
  loc_00E008E1: push eax
  loc_00E008E2: mov fs:[00000000h], esp
  loc_00E008E9: sub esp, 0000001Ch
  loc_00E008EC: push ebx
  loc_00E008ED: push esi
  loc_00E008EE: push edi
  loc_00E008EF: mov var_C, esp
  loc_00E008F2: mov var_8, 00401BC8h
  loc_00E008F9: xor edi, edi
  loc_00E008FB: mov var_4, edi
  loc_00E008FE: mov esi, Me
  loc_00E00901: push esi
  loc_00E00902: mov eax, [esi]
  loc_00E00904: call [eax+00000004h]
  loc_00E00907: mov edx, sAddress
  loc_00E0090A: lea ecx, var_18
  loc_00E0090D: mov var_18, edi
  loc_00E00910: mov var_1C, edi
  loc_00E00913: mov var_20, edi
  loc_00E00916: mov var_24, edi
  loc_00E00919: call [004011B0h] ; __vbaStrCopy
  loc_00E0091F: mov ecx, [esi]
  loc_00E00921: lea edx, var_24
  loc_00E00924: push edx
  loc_00E00925: push esi
  loc_00E00926: call [ecx+00000818h]
  loc_00E0092C: cmp eax, edi
  loc_00E0092E: jge 00E00942h
  loc_00E00930: push 00000818h
  loc_00E00935: push 006DD090h
  loc_00E0093A: push esi
  loc_00E0093B: push eax
  loc_00E0093C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E00942: mov eax, var_18
  loc_00E00945: mov esi, [004011ECh] ; __vbaStrToAnsi
  loc_00E0094B: push 00000001h
  loc_00E0094D: push edi
  loc_00E0094E: push edi
  loc_00E0094F: lea ecx, var_20
  loc_00E00952: push eax
  loc_00E00953: push ecx
  loc_00E00954: call __vbaStrToAnsi
  loc_00E00956: push eax
  loc_00E00957: lea edx, var_1C
  loc_00E0095A: push 006DF0E8h ; "open"
  loc_00E0095F: push edx
  loc_00E00960: call __vbaStrToAnsi
  loc_00E00962: push eax
  loc_00E00963: mov eax, var_24
  loc_00E00966: push eax
  loc_00E00967: call 006DDE0Ch ; ShellExecute(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v)
  loc_00E0096C: call [00401074h] ; __vbaSetSystemError
  loc_00E00972: mov ecx, var_20
  loc_00E00975: lea edx, var_18
  loc_00E00978: push ecx
  loc_00E00979: push edx
  loc_00E0097A: call [00401160h] ; __vbaStrToUnicode
  loc_00E00980: lea eax, var_20
  loc_00E00983: lea ecx, var_1C
  loc_00E00986: push eax
  loc_00E00987: push ecx
  loc_00E00988: push 00000002h
  loc_00E0098A: call [004011BCh] ; __vbaFreeStrList
  loc_00E00990: add esp, 0000000Ch
  loc_00E00993: push 00E009B8h
  loc_00E00998: jmp 00E009AEh
  loc_00E0099A: lea edx, var_20
  loc_00E0099D: lea eax, var_1C
  loc_00E009A0: push edx
  loc_00E009A1: push eax
  loc_00E009A2: push 00000002h
  loc_00E009A4: call [004011BCh] ; __vbaFreeStrList
  loc_00E009AA: add esp, 0000000Ch
  loc_00E009AD: ret
  loc_00E009AE: lea ecx, var_18
  loc_00E009B1: call [00401258h] ; __vbaFreeStr
  loc_00E009B7: ret
  loc_00E009B8: mov eax, Me
  loc_00E009BB: push eax
  loc_00E009BC: mov ecx, [eax]
  loc_00E009BE: call [ecx+00000008h]
  loc_00E009C1: mov eax, var_4
  loc_00E009C4: mov ecx, var_14
  loc_00E009C7: pop edi
  loc_00E009C8: pop esi
  loc_00E009C9: mov fs:[00000000h], ecx
  loc_00E009D0: pop ebx
  loc_00E009D1: mov esp, ebp
  loc_00E009D3: pop ebp
  loc_00E009D4: retn 0008h
End Sub

Public Sub about() 'E009E0
  loc_00E009E0: push ebp
  loc_00E009E1: mov ebp, esp
  loc_00E009E3: sub esp, 0000000Ch
  loc_00E009E6: push 00402836h ; __vbaExceptHandler
  loc_00E009EB: mov eax, fs:[00000000h]
  loc_00E009F1: push eax
  loc_00E009F2: mov fs:[00000000h], esp
  loc_00E009F9: sub esp, 00000090h
  loc_00E009FF: push ebx
  loc_00E00A00: push esi
  loc_00E00A01: push edi
  loc_00E00A02: mov var_C, esp
  loc_00E00A05: mov var_8, 00401BD8h
  loc_00E00A0C: xor esi, esi
  loc_00E00A0E: mov var_4, esi
  loc_00E00A11: mov eax, Me
  loc_00E00A14: push eax
  loc_00E00A15: mov ecx, [eax]
  loc_00E00A17: call [ecx+00000004h]
  loc_00E00A1A: mov ecx, 80020004h
  loc_00E00A1F: mov eax, 0000000Ah
  loc_00E00A24: mov var_64, ecx
  loc_00E00A27: mov var_54, ecx
  loc_00E00A2A: mov ebx, 00000008h
  loc_00E00A2F: mov var_5C, esi
  loc_00E00A32: mov var_6C, esi
  loc_00E00A35: mov var_7C, esi
  loc_00E00A38: lea edx, var_7C
  loc_00E00A3B: lea ecx, var_4C
  loc_00E00A3E: mov var_18, esi
  loc_00E00A41: mov var_1C, esi
  loc_00E00A44: mov var_20, esi
  loc_00E00A47: mov var_24, esi
  loc_00E00A4A: mov var_28, esi
  loc_00E00A4D: mov var_2C, esi
  loc_00E00A50: mov var_3C, esi
  loc_00E00A53: mov var_4C, esi
  loc_00E00A56: mov var_6C, eax
  loc_00E00A59: mov var_5C, eax
  loc_00E00A5C: mov var_74, 006DF218h ; "About"
  loc_00E00A63: mov var_7C, ebx
  loc_00E00A66: call [004011F0h] ; __vbaVarDup
  loc_00E00A6C: mov esi, [0040106Ch] ; __vbaStrCat
  loc_00E00A72: push 006DF0F8h ; "JCButton v 1.02"
  loc_00E00A77: push 006DF11Ch ; vbCrLf
  loc_00E00A7C: call __vbaStrCat
  loc_00E00A7E: mov edi, [00401228h] ; __vbaStrMove
  loc_00E00A84: mov edx, eax
  loc_00E00A86: lea ecx, var_18
  loc_00E00A89: call edi
  loc_00E00A8B: push eax
  loc_00E00A8C: push 006DF128h ; "Author: Juned S. Chhipa"
  loc_00E00A91: call __vbaStrCat
  loc_00E00A93: mov edx, eax
  loc_00E00A95: lea ecx, var_1C
  loc_00E00A98: call edi
  loc_00E00A9A: push eax
  loc_00E00A9B: push 006DF11Ch ; vbCrLf
  loc_00E00AA0: call __vbaStrCat
  loc_00E00AA2: mov edx, eax
  loc_00E00AA4: lea ecx, var_20
  loc_00E00AA7: call edi
  loc_00E00AA9: push eax
  loc_00E00AAA: push 006DF15Ch ; "Contact: juned.chhipa@yahoo.com"
  loc_00E00AAF: call __vbaStrCat
  loc_00E00AB1: mov edx, eax
  loc_00E00AB3: lea ecx, var_24
  loc_00E00AB6: call edi
  loc_00E00AB8: push eax
  loc_00E00AB9: push 006DF11Ch ; vbCrLf
  loc_00E00ABE: call __vbaStrCat
  loc_00E00AC0: mov edx, eax
  loc_00E00AC2: lea ecx, var_28
  loc_00E00AC5: call edi
  loc_00E00AC7: push eax
  loc_00E00AC8: push 006DF11Ch ; vbCrLf
  loc_00E00ACD: call __vbaStrCat
  loc_00E00ACF: mov edx, eax
  loc_00E00AD1: lea ecx, var_2C
  loc_00E00AD4: call edi
  loc_00E00AD6: push eax
  loc_00E00AD7: push 006DF1A0h ; "Copyright © 2008-2009 Juned Chhipa. All rights reserved."
  loc_00E00ADC: call __vbaStrCat
  loc_00E00ADE: mov var_34, eax
  loc_00E00AE1: lea edx, var_6C
  loc_00E00AE4: lea eax, var_5C
  loc_00E00AE7: push edx
  loc_00E00AE8: mov var_3C, ebx
  loc_00E00AEB: push eax
  loc_00E00AEC: lea ecx, var_4C
  loc_00E00AEF: lea edx, var_3C
  loc_00E00AF2: push ecx
  loc_00E00AF3: push 00000040h
  loc_00E00AF5: push edx
  loc_00E00AF6: call [004010A8h] ; rtcMsgBox
  loc_00E00AFC: lea eax, var_2C
  loc_00E00AFF: lea ecx, var_28
  loc_00E00B02: push eax
  loc_00E00B03: lea edx, var_24
  loc_00E00B06: push ecx
  loc_00E00B07: lea eax, var_20
  loc_00E00B0A: push edx
  loc_00E00B0B: lea ecx, var_1C
  loc_00E00B0E: push eax
  loc_00E00B0F: lea edx, var_18
  loc_00E00B12: push ecx
  loc_00E00B13: push edx
  loc_00E00B14: push 00000006h
  loc_00E00B16: call [004011BCh] ; __vbaFreeStrList
  loc_00E00B1C: lea eax, var_6C
  loc_00E00B1F: lea ecx, var_5C
  loc_00E00B22: push eax
  loc_00E00B23: lea edx, var_4C
  loc_00E00B26: push ecx
  loc_00E00B27: lea eax, var_3C
  loc_00E00B2A: push edx
  loc_00E00B2B: push eax
  loc_00E00B2C: push 00000004h
  loc_00E00B2E: call [00401038h] ; __vbaFreeVarList
  loc_00E00B34: add esp, 00000030h
  loc_00E00B37: push 00E00B7Bh
  loc_00E00B3C: jmp 00E00B7Ah
  loc_00E00B3E: lea ecx, var_2C
  loc_00E00B41: lea edx, var_28
  loc_00E00B44: push ecx
  loc_00E00B45: lea eax, var_24
  loc_00E00B48: push edx
  loc_00E00B49: lea ecx, var_20
  loc_00E00B4C: push eax
  loc_00E00B4D: lea edx, var_1C
  loc_00E00B50: push ecx
  loc_00E00B51: lea eax, var_18
  loc_00E00B54: push edx
  loc_00E00B55: push eax
  loc_00E00B56: push 00000006h
  loc_00E00B58: call [004011BCh] ; __vbaFreeStrList
  loc_00E00B5E: lea ecx, var_6C
  loc_00E00B61: lea edx, var_5C
  loc_00E00B64: push ecx
  loc_00E00B65: lea eax, var_4C
  loc_00E00B68: push edx
  loc_00E00B69: lea ecx, var_3C
  loc_00E00B6C: push eax
  loc_00E00B6D: push ecx
  loc_00E00B6E: push 00000004h
  loc_00E00B70: call [00401038h] ; __vbaFreeVarList
  loc_00E00B76: add esp, 00000030h
  loc_00E00B79: ret
  loc_00E00B7A: ret
  loc_00E00B7B: mov eax, Me
  loc_00E00B7E: push eax
  loc_00E00B7F: mov edx, [eax]
  loc_00E00B81: call [edx+00000008h]
  loc_00E00B84: mov eax, var_4
  loc_00E00B87: mov ecx, var_14
  loc_00E00B8A: pop edi
  loc_00E00B8B: pop esi
  loc_00E00B8C: mov fs:[00000000h], ecx
  loc_00E00B93: pop ebx
  loc_00E00B94: mov esp, ebp
  loc_00E00B96: pop ebp
  loc_00E00B97: retn 0004h
End Sub

Public Property Get BackColor(arg_C) 'E00BA0
  loc_00E00BA0: push ebp
  loc_00E00BA1: mov ebp, esp
  loc_00E00BA3: sub esp, 0000000Ch
  loc_00E00BA6: push 00402836h ; __vbaExceptHandler
  loc_00E00BAB: mov eax, fs:[00000000h]
  loc_00E00BB1: push eax
  loc_00E00BB2: mov fs:[00000000h], esp
  loc_00E00BB9: sub esp, 0000000Ch
  loc_00E00BBC: push ebx
  loc_00E00BBD: push esi
  loc_00E00BBE: push edi
  loc_00E00BBF: mov var_C, esp
  loc_00E00BC2: mov var_8, 00401BE8h
  loc_00E00BC9: xor edi, edi
  loc_00E00BCB: mov var_4, edi
  loc_00E00BCE: mov esi, Me
  loc_00E00BD1: push esi
  loc_00E00BD2: mov eax, [esi]
  loc_00E00BD4: call [eax+00000004h]
  loc_00E00BD7: mov ecx, [esi+00000088h]
  loc_00E00BDD: mov var_18, edi
  loc_00E00BE0: mov var_18, ecx
  loc_00E00BE3: mov eax, Me
  loc_00E00BE6: push eax
  loc_00E00BE7: mov edx, [eax]
  loc_00E00BE9: call [edx+00000008h]
  loc_00E00BEC: mov eax, arg_C
  loc_00E00BEF: mov ecx, var_18
  loc_00E00BF2: mov [eax], ecx
  loc_00E00BF4: mov eax, var_4
  loc_00E00BF7: mov ecx, var_14
  loc_00E00BFA: pop edi
  loc_00E00BFB: pop esi
  loc_00E00BFC: mov fs:[00000000h], ecx
  loc_00E00C03: pop ebx
  loc_00E00C04: mov esp, ebp
  loc_00E00C06: pop ebp
  loc_00E00C07: retn 0008h
End Sub

Public Property Let BackColor(New_BackColor) 'E00C10
  loc_00E00C10: push ebp
  loc_00E00C11: mov ebp, esp
  loc_00E00C13: sub esp, 0000000Ch
  loc_00E00C16: push 00402836h ; __vbaExceptHandler
  loc_00E00C1B: mov eax, fs:[00000000h]
  loc_00E00C21: push eax
  loc_00E00C22: mov fs:[00000000h], esp
  loc_00E00C29: sub esp, 0000001Ch
  loc_00E00C2C: push ebx
  loc_00E00C2D: push esi
  loc_00E00C2E: push edi
  loc_00E00C2F: mov var_C, esp
  loc_00E00C32: mov var_8, 00401BF0h
  loc_00E00C39: mov var_4, 00000000h
  loc_00E00C40: mov esi, Me
  loc_00E00C43: push esi
  loc_00E00C44: mov eax, [esi]
  loc_00E00C46: call [eax+00000004h]
  loc_00E00C49: mov eax, [esi+00000044h]
  loc_00E00C4C: mov ecx, New_BackColor
  loc_00E00C4F: cmp eax, 00000004h
  loc_00E00C52: mov [esi+00000088h], ecx
  loc_00E00C58: jz 00E00C64h
  loc_00E00C5A: mov [esi+000000CCh], 00000003h
  loc_00E00C64: mov edx, [esi]
  loc_00E00C66: push esi
  loc_00E00C67: call [edx+00000910h]
  loc_00E00C6D: sub esp, 00000010h
  loc_00E00C70: mov ecx, 00000008h
  loc_00E00C75: mov edi, esp
  loc_00E00C77: mov edx, [esi]
  loc_00E00C79: mov eax, 006DEACCh ; "BackColor"
  loc_00E00C7E: push esi
  loc_00E00C7F: mov [edi], ecx
  loc_00E00C81: mov ecx, var_20
  loc_00E00C84: mov [edi+00000004h], ecx
  loc_00E00C87: mov [edi+00000008h], eax
  loc_00E00C8A: mov eax, var_18
  loc_00E00C8D: mov [edi+0000000Ch], eax
  loc_00E00C90: call [edx+00000390h]
  loc_00E00C96: test eax, eax
  loc_00E00C98: fnclex
  loc_00E00C9A: jge 00E00CAEh
  loc_00E00C9C: push 00000390h
  loc_00E00CA1: push 006DCB00h
  loc_00E00CA6: push esi
  loc_00E00CA7: push eax
  loc_00E00CA8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E00CAE: mov eax, Me
  loc_00E00CB1: push eax
  loc_00E00CB2: mov ecx, [eax]
  loc_00E00CB4: call [ecx+00000008h]
  loc_00E00CB7: mov eax, var_4
  loc_00E00CBA: mov ecx, var_14
  loc_00E00CBD: pop edi
  loc_00E00CBE: pop esi
  loc_00E00CBF: mov fs:[00000000h], ecx
  loc_00E00CC6: pop ebx
  loc_00E00CC7: mov esp, ebp
  loc_00E00CC9: pop ebp
  loc_00E00CCA: retn 0008h
End Sub

Public Property Get ButtonStyle(arg_C) 'E00CD0
  loc_00E00CD0: push ebp
  loc_00E00CD1: mov ebp, esp
  loc_00E00CD3: sub esp, 0000000Ch
  loc_00E00CD6: push 00402836h ; __vbaExceptHandler
  loc_00E00CDB: mov eax, fs:[00000000h]
  loc_00E00CE1: push eax
  loc_00E00CE2: mov fs:[00000000h], esp
  loc_00E00CE9: sub esp, 0000000Ch
  loc_00E00CEC: push ebx
  loc_00E00CED: push esi
  loc_00E00CEE: push edi
  loc_00E00CEF: mov var_C, esp
  loc_00E00CF2: mov var_8, 00401BF8h
  loc_00E00CF9: xor edi, edi
  loc_00E00CFB: mov var_4, edi
  loc_00E00CFE: mov esi, Me
  loc_00E00D01: push esi
  loc_00E00D02: mov eax, [esi]
  loc_00E00D04: call [eax+00000004h]
  loc_00E00D07: mov ecx, [esi+00000044h]
  loc_00E00D0A: mov var_18, edi
  loc_00E00D0D: mov var_18, ecx
  loc_00E00D10: mov eax, Me
  loc_00E00D13: push eax
  loc_00E00D14: mov edx, [eax]
  loc_00E00D16: call [edx+00000008h]
  loc_00E00D19: mov eax, arg_C
  loc_00E00D1C: mov ecx, var_18
  loc_00E00D1F: mov [eax], ecx
  loc_00E00D21: mov eax, var_4
  loc_00E00D24: mov ecx, var_14
  loc_00E00D27: pop edi
  loc_00E00D28: pop esi
  loc_00E00D29: mov fs:[00000000h], ecx
  loc_00E00D30: pop ebx
  loc_00E00D31: mov esp, ebp
  loc_00E00D33: pop ebp
  loc_00E00D34: retn 0008h
End Sub

Public Property Let ButtonStyle(New_ButtonStyle) 'E00D40
  loc_00E00D40: push ebp
  loc_00E00D41: mov ebp, esp
  loc_00E00D43: sub esp, 0000000Ch
  loc_00E00D46: push 00402836h ; __vbaExceptHandler
  loc_00E00D4B: mov eax, fs:[00000000h]
  loc_00E00D51: push eax
  loc_00E00D52: mov fs:[00000000h], esp
  loc_00E00D59: sub esp, 0000001Ch
  loc_00E00D5C: push ebx
  loc_00E00D5D: push esi
  loc_00E00D5E: push edi
  loc_00E00D5F: mov var_C, esp
  loc_00E00D62: mov var_8, 00401C00h
  loc_00E00D69: mov var_4, 00000000h
  loc_00E00D70: mov esi, Me
  loc_00E00D73: push esi
  loc_00E00D74: mov eax, [esi]
  loc_00E00D76: call [eax+00000004h]
  loc_00E00D79: mov ecx, New_ButtonStyle
  loc_00E00D7C: mov edx, [esi]
  loc_00E00D7E: push esi
  loc_00E00D7F: mov [esi+00000044h], ecx
  loc_00E00D82: call [edx+000009B4h]
  loc_00E00D88: mov eax, [esi]
  loc_00E00D8A: push esi
  loc_00E00D8B: call [eax+000009B8h]
  loc_00E00D91: mov ecx, [esi]
  loc_00E00D93: push esi
  loc_00E00D94: call [ecx+00000914h]
  loc_00E00D9A: mov edx, [esi]
  loc_00E00D9C: push esi
  loc_00E00D9D: call [edx+00000910h]
  loc_00E00DA3: sub esp, 00000010h
  loc_00E00DA6: mov ecx, 00000008h
  loc_00E00DAB: mov edi, esp
  loc_00E00DAD: mov edx, [esi]
  loc_00E00DAF: mov eax, 006DEB60h ; "ButtonStyle"
  loc_00E00DB4: push esi
  loc_00E00DB5: mov [edi], ecx
  loc_00E00DB7: mov ecx, var_20
  loc_00E00DBA: mov [edi+00000004h], ecx
  loc_00E00DBD: mov [edi+00000008h], eax
  loc_00E00DC0: mov eax, var_18
  loc_00E00DC3: mov [edi+0000000Ch], eax
  loc_00E00DC6: call [edx+00000390h]
  loc_00E00DCC: test eax, eax
  loc_00E00DCE: fnclex
  loc_00E00DD0: jge 00E00DE4h
  loc_00E00DD2: push 00000390h
  loc_00E00DD7: push 006DCB00h
  loc_00E00DDC: push esi
  loc_00E00DDD: push eax
  loc_00E00DDE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E00DE4: mov eax, Me
  loc_00E00DE7: push eax
  loc_00E00DE8: mov ecx, [eax]
  loc_00E00DEA: call [ecx+00000008h]
  loc_00E00DED: mov eax, var_4
  loc_00E00DF0: mov ecx, var_14
  loc_00E00DF3: pop edi
  loc_00E00DF4: pop esi
  loc_00E00DF5: mov fs:[00000000h], ecx
  loc_00E00DFC: pop ebx
  loc_00E00DFD: mov esp, ebp
  loc_00E00DFF: pop ebp
  loc_00E00E00: retn 0008h
End Sub

Public Property Get Caption(arg_C) 'E00E10
  loc_00E00E10: push ebp
  loc_00E00E11: mov ebp, esp
  loc_00E00E13: sub esp, 0000000Ch
  loc_00E00E16: push 00402836h ; __vbaExceptHandler
  loc_00E00E1B: mov eax, fs:[00000000h]
  loc_00E00E21: push eax
  loc_00E00E22: mov fs:[00000000h], esp
  loc_00E00E29: sub esp, 0000000Ch
  loc_00E00E2C: push ebx
  loc_00E00E2D: push esi
  loc_00E00E2E: push edi
  loc_00E00E2F: mov var_C, esp
  loc_00E00E32: mov var_8, 00401C08h
  loc_00E00E39: xor edi, edi
  loc_00E00E3B: mov var_4, edi
  loc_00E00E3E: mov esi, Me
  loc_00E00E41: push esi
  loc_00E00E42: mov eax, [esi]
  loc_00E00E44: call [eax+00000004h]
  loc_00E00E47: mov ecx, arg_C
  loc_00E00E4A: mov var_18, edi
  loc_00E00E4D: mov [ecx], edi
  loc_00E00E4F: mov edx, [esi+00000080h]
  loc_00E00E55: lea ecx, var_18
  loc_00E00E58: call [004011B0h] ; __vbaStrCopy
  loc_00E00E5E: push 00E00E70h
  loc_00E00E63: jmp 00E00E6Fh
  loc_00E00E65: lea ecx, var_18
  loc_00E00E68: call [00401258h] ; __vbaFreeStr
  loc_00E00E6E: ret
  loc_00E00E6F: ret
  loc_00E00E70: mov eax, Me
  loc_00E00E73: push eax
  loc_00E00E74: mov edx, [eax]
  loc_00E00E76: call [edx+00000008h]
  loc_00E00E79: mov eax, arg_C
  loc_00E00E7C: mov ecx, var_18
  loc_00E00E7F: mov [eax], ecx
  loc_00E00E81: mov eax, var_4
  loc_00E00E84: mov ecx, var_14
  loc_00E00E87: pop edi
  loc_00E00E88: pop esi
  loc_00E00E89: mov fs:[00000000h], ecx
  loc_00E00E90: pop ebx
  loc_00E00E91: mov esp, ebp
  loc_00E00E93: pop ebp
  loc_00E00E94: retn 0008h
End Sub

Public Property Let Caption(New_Caption) 'E00EA0
  loc_00E00EA0: push ebp
  loc_00E00EA1: mov ebp, esp
  loc_00E00EA3: sub esp, 0000000Ch
  loc_00E00EA6: push 00402836h ; __vbaExceptHandler
  loc_00E00EAB: mov eax, fs:[00000000h]
  loc_00E00EB1: push eax
  loc_00E00EB2: mov fs:[00000000h], esp
  loc_00E00EB9: sub esp, 00000020h
  loc_00E00EBC: push ebx
  loc_00E00EBD: push esi
  loc_00E00EBE: push edi
  loc_00E00EBF: mov var_C, esp
  loc_00E00EC2: mov var_8, 00401C18h
  loc_00E00EC9: xor ebx, ebx
  loc_00E00ECB: mov var_4, ebx
  loc_00E00ECE: mov esi, Me
  loc_00E00ED1: push esi
  loc_00E00ED2: mov eax, [esi]
  loc_00E00ED4: call [eax+00000004h]
  loc_00E00ED7: mov edx, New_Caption
  loc_00E00EDA: mov edi, [004011B0h] ; __vbaStrCopy
  loc_00E00EE0: lea ecx, var_18
  loc_00E00EE3: mov var_18, ebx
  loc_00E00EE6: call edi
  loc_00E00EE8: mov edx, var_18
  loc_00E00EEB: lea ecx, [esi+00000080h]
  loc_00E00EF1: call edi
  loc_00E00EF3: mov ecx, [esi]
  loc_00E00EF5: push esi
  loc_00E00EF6: call [ecx+0000093Ch]
  loc_00E00EFC: mov edx, [esi]
  loc_00E00EFE: push esi
  loc_00E00EFF: call [edx+00000910h]
  loc_00E00F05: sub esp, 00000010h
  loc_00E00F08: mov ecx, 00000008h
  loc_00E00F0D: mov edi, esp
  loc_00E00F0F: mov edx, [esi]
  loc_00E00F11: mov eax, 006DEBD0h ; "Caption"
  loc_00E00F16: push esi
  loc_00E00F17: mov [edi], ecx
  loc_00E00F19: mov ecx, var_24
  loc_00E00F1C: mov [edi+00000004h], ecx
  loc_00E00F1F: mov [edi+00000008h], eax
  loc_00E00F22: mov eax, var_1C
  loc_00E00F25: mov [edi+0000000Ch], eax
  loc_00E00F28: call [edx+00000390h]
  loc_00E00F2E: cmp eax, ebx
  loc_00E00F30: fnclex
  loc_00E00F32: jge 00E00F46h
  loc_00E00F34: push 00000390h
  loc_00E00F39: push 006DCB00h
  loc_00E00F3E: push esi
  loc_00E00F3F: push eax
  loc_00E00F40: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E00F46: push 00E00F55h
  loc_00E00F4B: lea ecx, var_18
  loc_00E00F4E: call [00401258h] ; __vbaFreeStr
  loc_00E00F54: ret
  loc_00E00F55: mov eax, Me
  loc_00E00F58: push eax
  loc_00E00F59: mov ecx, [eax]
  loc_00E00F5B: call [ecx+00000008h]
  loc_00E00F5E: mov eax, var_4
  loc_00E00F61: mov ecx, var_14
  loc_00E00F64: pop edi
  loc_00E00F65: pop esi
  loc_00E00F66: mov fs:[00000000h], ecx
  loc_00E00F6D: pop ebx
  loc_00E00F6E: mov esp, ebp
  loc_00E00F70: pop ebp
  loc_00E00F71: retn 0008h
End Sub

Public Property Get CaptionAlign(arg_C) 'E00F80
  loc_00E00F80: push ebp
  loc_00E00F81: mov ebp, esp
  loc_00E00F83: sub esp, 0000000Ch
  loc_00E00F86: push 00402836h ; __vbaExceptHandler
  loc_00E00F8B: mov eax, fs:[00000000h]
  loc_00E00F91: push eax
  loc_00E00F92: mov fs:[00000000h], esp
  loc_00E00F99: sub esp, 0000000Ch
  loc_00E00F9C: push ebx
  loc_00E00F9D: push esi
  loc_00E00F9E: push edi
  loc_00E00F9F: mov var_C, esp
  loc_00E00FA2: mov var_8, 00401C28h
  loc_00E00FA9: xor edi, edi
  loc_00E00FAB: mov var_4, edi
  loc_00E00FAE: mov esi, Me
  loc_00E00FB1: push esi
  loc_00E00FB2: mov eax, [esi]
  loc_00E00FB4: call [eax+00000004h]
  loc_00E00FB7: mov ecx, [esi+00000084h]
  loc_00E00FBD: mov var_18, edi
  loc_00E00FC0: mov var_18, ecx
  loc_00E00FC3: mov eax, Me
  loc_00E00FC6: push eax
  loc_00E00FC7: mov edx, [eax]
  loc_00E00FC9: call [edx+00000008h]
  loc_00E00FCC: mov eax, arg_C
  loc_00E00FCF: mov ecx, var_18
  loc_00E00FD2: mov [eax], ecx
  loc_00E00FD4: mov eax, var_4
  loc_00E00FD7: mov ecx, var_14
  loc_00E00FDA: pop edi
  loc_00E00FDB: pop esi
  loc_00E00FDC: mov fs:[00000000h], ecx
  loc_00E00FE3: pop ebx
  loc_00E00FE4: mov esp, ebp
  loc_00E00FE6: pop ebp
  loc_00E00FE7: retn 0008h
End Sub

Public Property Let CaptionAlign(New_CaptionAlign) 'E00FF0
  loc_00E00FF0: push ebp
  loc_00E00FF1: mov ebp, esp
  loc_00E00FF3: sub esp, 0000000Ch
  loc_00E00FF6: push 00402836h ; __vbaExceptHandler
  loc_00E00FFB: mov eax, fs:[00000000h]
  loc_00E01001: push eax
  loc_00E01002: mov fs:[00000000h], esp
  loc_00E01009: sub esp, 0000001Ch
  loc_00E0100C: push ebx
  loc_00E0100D: push esi
  loc_00E0100E: push edi
  loc_00E0100F: mov var_C, esp
  loc_00E01012: mov var_8, 00401C30h
  loc_00E01019: mov var_4, 00000000h
  loc_00E01020: mov esi, Me
  loc_00E01023: push esi
  loc_00E01024: mov eax, [esi]
  loc_00E01026: call [eax+00000004h]
  loc_00E01029: mov ecx, New_CaptionAlign
  loc_00E0102C: mov edx, [esi]
  loc_00E0102E: push esi
  loc_00E0102F: mov [esi+00000084h], ecx
  loc_00E01035: call [edx+00000910h]
  loc_00E0103B: sub esp, 00000010h
  loc_00E0103E: mov ecx, 00000008h
  loc_00E01043: mov edi, esp
  loc_00E01045: mov edx, [esi]
  loc_00E01047: mov eax, 006DEE74h ; "CaptionAlign"
  loc_00E0104C: push esi
  loc_00E0104D: mov [edi], ecx
  loc_00E0104F: mov ecx, var_20
  loc_00E01052: mov [edi+00000004h], ecx
  loc_00E01055: mov [edi+00000008h], eax
  loc_00E01058: mov eax, var_18
  loc_00E0105B: mov [edi+0000000Ch], eax
  loc_00E0105E: call [edx+00000390h]
  loc_00E01064: test eax, eax
  loc_00E01066: fnclex
  loc_00E01068: jge 00E0107Ch
  loc_00E0106A: push 00000390h
  loc_00E0106F: push 006DCB00h
  loc_00E01074: push esi
  loc_00E01075: push eax
  loc_00E01076: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0107C: mov eax, Me
  loc_00E0107F: push eax
  loc_00E01080: mov ecx, [eax]
  loc_00E01082: call [ecx+00000008h]
  loc_00E01085: mov eax, var_4
  loc_00E01088: mov ecx, var_14
  loc_00E0108B: pop edi
  loc_00E0108C: pop esi
  loc_00E0108D: mov fs:[00000000h], ecx
  loc_00E01094: pop ebx
  loc_00E01095: mov esp, ebp
  loc_00E01097: pop ebp
  loc_00E01098: retn 0008h
End Sub

Public Property Get DropDownSymbol(arg_C) 'E010A0
  loc_00E010A0: push ebp
  loc_00E010A1: mov ebp, esp
  loc_00E010A3: sub esp, 0000000Ch
  loc_00E010A6: push 00402836h ; __vbaExceptHandler
  loc_00E010AB: mov eax, fs:[00000000h]
  loc_00E010B1: push eax
  loc_00E010B2: mov fs:[00000000h], esp
  loc_00E010B9: sub esp, 0000000Ch
  loc_00E010BC: push ebx
  loc_00E010BD: push esi
  loc_00E010BE: push edi
  loc_00E010BF: mov var_C, esp
  loc_00E010C2: mov var_8, 00401C38h
  loc_00E010C9: xor edi, edi
  loc_00E010CB: mov var_4, edi
  loc_00E010CE: mov esi, Me
  loc_00E010D1: push esi
  loc_00E010D2: mov eax, [esi]
  loc_00E010D4: call [eax+00000004h]
  loc_00E010D7: mov ecx, [esi+0000005Ch]
  loc_00E010DA: mov var_18, edi
  loc_00E010DD: mov var_18, ecx
  loc_00E010E0: mov eax, Me
  loc_00E010E3: push eax
  loc_00E010E4: mov edx, [eax]
  loc_00E010E6: call [edx+00000008h]
  loc_00E010E9: mov eax, arg_C
  loc_00E010EC: mov ecx, var_18
  loc_00E010EF: mov [eax], ecx
  loc_00E010F1: mov eax, var_4
  loc_00E010F4: mov ecx, var_14
  loc_00E010F7: pop edi
  loc_00E010F8: pop esi
  loc_00E010F9: mov fs:[00000000h], ecx
  loc_00E01100: pop ebx
  loc_00E01101: mov esp, ebp
  loc_00E01103: pop ebp
  loc_00E01104: retn 0008h
End Sub

Public Property Let DropDownSymbol(New_Align) 'E01110
  loc_00E01110: push ebp
  loc_00E01111: mov ebp, esp
  loc_00E01113: sub esp, 0000000Ch
  loc_00E01116: push 00402836h ; __vbaExceptHandler
  loc_00E0111B: mov eax, fs:[00000000h]
  loc_00E01121: push eax
  loc_00E01122: mov fs:[00000000h], esp
  loc_00E01129: sub esp, 0000001Ch
  loc_00E0112C: push ebx
  loc_00E0112D: push esi
  loc_00E0112E: push edi
  loc_00E0112F: mov var_C, esp
  loc_00E01132: mov var_8, 00401C40h
  loc_00E01139: mov var_4, 00000000h
  loc_00E01140: mov esi, Me
  loc_00E01143: push esi
  loc_00E01144: mov eax, [esi]
  loc_00E01146: call [eax+00000004h]
  loc_00E01149: mov ecx, New_Align
  loc_00E0114C: mov edx, [esi]
  loc_00E0114E: push esi
  loc_00E0114F: mov [esi+0000005Ch], ecx
  loc_00E01152: call [edx+00000910h]
  loc_00E01158: sub esp, 00000010h
  loc_00E0115B: mov ecx, 00000008h
  loc_00E01160: mov edi, esp
  loc_00E01162: mov edx, [esi]
  loc_00E01164: mov eax, 006DEFA8h ; "DropDownSymbol"
  loc_00E01169: push esi
  loc_00E0116A: mov [edi], ecx
  loc_00E0116C: mov ecx, var_20
  loc_00E0116F: mov [edi+00000004h], ecx
  loc_00E01172: mov [edi+00000008h], eax
  loc_00E01175: mov eax, var_18
  loc_00E01178: mov [edi+0000000Ch], eax
  loc_00E0117B: call [edx+00000390h]
  loc_00E01181: test eax, eax
  loc_00E01183: fnclex
  loc_00E01185: jge 00E01199h
  loc_00E01187: push 00000390h
  loc_00E0118C: push 006DCB00h
  loc_00E01191: push esi
  loc_00E01192: push eax
  loc_00E01193: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01199: mov eax, Me
  loc_00E0119C: push eax
  loc_00E0119D: mov ecx, [eax]
  loc_00E0119F: call [ecx+00000008h]
  loc_00E011A2: mov eax, var_4
  loc_00E011A5: mov ecx, var_14
  loc_00E011A8: pop edi
  loc_00E011A9: pop esi
  loc_00E011AA: mov fs:[00000000h], ecx
  loc_00E011B1: pop ebx
  loc_00E011B2: mov esp, ebp
  loc_00E011B4: pop ebp
  loc_00E011B5: retn 0008h
End Sub

Public Property Get DropDownSeparator(arg_C) 'E011C0
  loc_00E011C0: push ebp
  loc_00E011C1: mov ebp, esp
  loc_00E011C3: sub esp, 0000000Ch
  loc_00E011C6: push 00402836h ; __vbaExceptHandler
  loc_00E011CB: mov eax, fs:[00000000h]
  loc_00E011D1: push eax
  loc_00E011D2: mov fs:[00000000h], esp
  loc_00E011D9: sub esp, 0000000Ch
  loc_00E011DC: push ebx
  loc_00E011DD: push esi
  loc_00E011DE: push edi
  loc_00E011DF: mov var_C, esp
  loc_00E011E2: mov var_8, 00401C48h
  loc_00E011E9: xor edi, edi
  loc_00E011EB: mov var_4, edi
  loc_00E011EE: mov esi, Me
  loc_00E011F1: push esi
  loc_00E011F2: mov eax, [esi]
  loc_00E011F4: call [eax+00000004h]
  loc_00E011F7: mov cx, [esi+00000060h]
  loc_00E011FB: mov var_18, edi
  loc_00E011FE: mov var_18, ecx
  loc_00E01201: mov eax, Me
  loc_00E01204: push eax
  loc_00E01205: mov edx, [eax]
  loc_00E01207: call [edx+00000008h]
  loc_00E0120A: mov eax, arg_C
  loc_00E0120D: mov cx, var_18
  loc_00E01211: mov [eax], cx
  loc_00E01214: mov eax, var_4
  loc_00E01217: mov ecx, var_14
  loc_00E0121A: pop edi
  loc_00E0121B: pop esi
  loc_00E0121C: mov fs:[00000000h], ecx
  loc_00E01223: pop ebx
  loc_00E01224: mov esp, ebp
  loc_00E01226: pop ebp
  loc_00E01227: retn 0008h
End Sub

Public Property Let DropDownSeparator(New_Value) 'E01230
  loc_00E01230: push ebp
  loc_00E01231: mov ebp, esp
  loc_00E01233: sub esp, 0000000Ch
  loc_00E01236: push 00402836h ; __vbaExceptHandler
  loc_00E0123B: mov eax, fs:[00000000h]
  loc_00E01241: push eax
  loc_00E01242: mov fs:[00000000h], esp
  loc_00E01249: sub esp, 0000001Ch
  loc_00E0124C: push ebx
  loc_00E0124D: push esi
  loc_00E0124E: push edi
  loc_00E0124F: mov var_C, esp
  loc_00E01252: mov var_8, 00401C50h
  loc_00E01259: mov var_4, 00000000h
  loc_00E01260: mov esi, Me
  loc_00E01263: push esi
  loc_00E01264: mov eax, [esi]
  loc_00E01266: call [eax+00000004h]
  loc_00E01269: mov cx, New_Value
  loc_00E0126D: mov edx, [esi]
  loc_00E0126F: push esi
  loc_00E01270: mov [esi+00000060h], cx
  loc_00E01274: call [edx+00000910h]
  loc_00E0127A: sub esp, 00000010h
  loc_00E0127D: mov ecx, 00000008h
  loc_00E01282: mov edi, esp
  loc_00E01284: mov edx, [esi]
  loc_00E01286: mov eax, 006DEED0h ; "DropDownSeparator"
  loc_00E0128B: push esi
  loc_00E0128C: mov [edi], ecx
  loc_00E0128E: mov ecx, var_20
  loc_00E01291: mov [edi+00000004h], ecx
  loc_00E01294: mov [edi+00000008h], eax
  loc_00E01297: mov eax, var_18
  loc_00E0129A: mov [edi+0000000Ch], eax
  loc_00E0129D: call [edx+00000390h]
  loc_00E012A3: test eax, eax
  loc_00E012A5: fnclex
  loc_00E012A7: jge 00E012BBh
  loc_00E012A9: push 00000390h
  loc_00E012AE: push 006DCB00h
  loc_00E012B3: push esi
  loc_00E012B4: push eax
  loc_00E012B5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E012BB: mov eax, Me
  loc_00E012BE: push eax
  loc_00E012BF: mov ecx, [eax]
  loc_00E012C1: call [ecx+00000008h]
  loc_00E012C4: mov eax, var_4
  loc_00E012C7: mov ecx, var_14
  loc_00E012CA: pop edi
  loc_00E012CB: pop esi
  loc_00E012CC: mov fs:[00000000h], ecx
  loc_00E012D3: pop ebx
  loc_00E012D4: mov esp, ebp
  loc_00E012D6: pop ebp
  loc_00E012D7: retn 0008h
End Sub

Public Property Get DisabledPictureMode(arg_C) 'E012E0
  loc_00E012E0: push ebp
  loc_00E012E1: mov ebp, esp
  loc_00E012E3: sub esp, 0000000Ch
  loc_00E012E6: push 00402836h ; __vbaExceptHandler
  loc_00E012EB: mov eax, fs:[00000000h]
  loc_00E012F1: push eax
  loc_00E012F2: mov fs:[00000000h], esp
  loc_00E012F9: sub esp, 0000000Ch
  loc_00E012FC: push ebx
  loc_00E012FD: push esi
  loc_00E012FE: push edi
  loc_00E012FF: mov var_C, esp
  loc_00E01302: mov var_8, 00401C58h
  loc_00E01309: xor edi, edi
  loc_00E0130B: mov var_4, edi
  loc_00E0130E: mov esi, Me
  loc_00E01311: push esi
  loc_00E01312: mov eax, [esi]
  loc_00E01314: call [eax+00000004h]
  loc_00E01317: mov ecx, [esi+0000015Ch]
  loc_00E0131D: mov var_18, edi
  loc_00E01320: mov var_18, ecx
  loc_00E01323: mov eax, Me
  loc_00E01326: push eax
  loc_00E01327: mov edx, [eax]
  loc_00E01329: call [edx+00000008h]
  loc_00E0132C: mov eax, arg_C
  loc_00E0132F: mov ecx, var_18
  loc_00E01332: mov [eax], ecx
  loc_00E01334: mov eax, var_4
  loc_00E01337: mov ecx, var_14
  loc_00E0133A: pop edi
  loc_00E0133B: pop esi
  loc_00E0133C: mov fs:[00000000h], ecx
  loc_00E01343: pop ebx
  loc_00E01344: mov esp, ebp
  loc_00E01346: pop ebp
  loc_00E01347: retn 0008h
End Sub

Public Property Let DisabledPictureMode(New_mode) 'E01350
  loc_00E01350: push ebp
  loc_00E01351: mov ebp, esp
  loc_00E01353: sub esp, 0000000Ch
  loc_00E01356: push 00402836h ; __vbaExceptHandler
  loc_00E0135B: mov eax, fs:[00000000h]
  loc_00E01361: push eax
  loc_00E01362: mov fs:[00000000h], esp
  loc_00E01369: sub esp, 0000001Ch
  loc_00E0136C: push ebx
  loc_00E0136D: push esi
  loc_00E0136E: push edi
  loc_00E0136F: mov var_C, esp
  loc_00E01372: mov var_8, 00401C60h
  loc_00E01379: mov var_4, 00000000h
  loc_00E01380: mov esi, Me
  loc_00E01383: push esi
  loc_00E01384: mov eax, [esi]
  loc_00E01386: call [eax+00000004h]
  loc_00E01389: mov ecx, New_mode
  loc_00E0138C: mov edx, [esi]
  loc_00E0138E: push esi
  loc_00E0138F: mov [esi+0000015Ch], ecx
  loc_00E01395: call [edx+00000910h]
  loc_00E0139B: sub esp, 00000010h
  loc_00E0139E: mov ecx, 00000008h
  loc_00E013A3: mov edi, esp
  loc_00E013A5: mov edx, [esi]
  loc_00E013A7: mov eax, 006DED38h ; "DisabledPictureMode"
  loc_00E013AC: push esi
  loc_00E013AD: mov [edi], ecx
  loc_00E013AF: mov ecx, var_20
  loc_00E013B2: mov [edi+00000004h], ecx
  loc_00E013B5: mov [edi+00000008h], eax
  loc_00E013B8: mov eax, var_18
  loc_00E013BB: mov [edi+0000000Ch], eax
  loc_00E013BE: call [edx+00000390h]
  loc_00E013C4: test eax, eax
  loc_00E013C6: fnclex
  loc_00E013C8: jge 00E013DCh
  loc_00E013CA: push 00000390h
  loc_00E013CF: push 006DCB00h
  loc_00E013D4: push esi
  loc_00E013D5: push eax
  loc_00E013D6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E013DC: mov eax, Me
  loc_00E013DF: push eax
  loc_00E013E0: mov ecx, [eax]
  loc_00E013E2: call [ecx+00000008h]
  loc_00E013E5: mov eax, var_4
  loc_00E013E8: mov ecx, var_14
  loc_00E013EB: pop edi
  loc_00E013EC: pop esi
  loc_00E013ED: mov fs:[00000000h], ecx
  loc_00E013F4: pop ebx
  loc_00E013F5: mov esp, ebp
  loc_00E013F7: pop ebp
  loc_00E013F8: retn 0008h
End Sub

Public Property Get Enabled(arg_C) 'E01400
  loc_00E01400: push ebp
  loc_00E01401: mov ebp, esp
  loc_00E01403: sub esp, 0000000Ch
  loc_00E01406: push 00402836h ; __vbaExceptHandler
  loc_00E0140B: mov eax, fs:[00000000h]
  loc_00E01411: push eax
  loc_00E01412: mov fs:[00000000h], esp
  loc_00E01419: sub esp, 0000000Ch
  loc_00E0141C: push ebx
  loc_00E0141D: push esi
  loc_00E0141E: push edi
  loc_00E0141F: mov var_C, esp
  loc_00E01422: mov var_8, 00401C68h
  loc_00E01429: xor edi, edi
  loc_00E0142B: mov var_4, edi
  loc_00E0142E: mov esi, Me
  loc_00E01431: push esi
  loc_00E01432: mov eax, [esi]
  loc_00E01434: call [eax+00000004h]
  loc_00E01437: mov cx, [esi+0000007Ch]
  loc_00E0143B: mov var_18, edi
  loc_00E0143E: mov var_18, ecx
  loc_00E01441: mov eax, Me
  loc_00E01444: push eax
  loc_00E01445: mov edx, [eax]
  loc_00E01447: call [edx+00000008h]
  loc_00E0144A: mov eax, arg_C
  loc_00E0144D: mov cx, var_18
  loc_00E01451: mov [eax], cx
  loc_00E01454: mov eax, var_4
  loc_00E01457: mov ecx, var_14
  loc_00E0145A: pop edi
  loc_00E0145B: pop esi
  loc_00E0145C: mov fs:[00000000h], ecx
  loc_00E01463: pop ebx
  loc_00E01464: mov esp, ebp
  loc_00E01466: pop ebp
  loc_00E01467: retn 0008h
End Sub

Public Property Let Enabled(New_Enabled) 'E01470
  loc_00E01470: push ebp
  loc_00E01471: mov ebp, esp
  loc_00E01473: sub esp, 0000000Ch
  loc_00E01476: push 00402836h ; __vbaExceptHandler
  loc_00E0147B: mov eax, fs:[00000000h]
  loc_00E01481: push eax
  loc_00E01482: mov fs:[00000000h], esp
  loc_00E01489: sub esp, 0000001Ch
  loc_00E0148C: push ebx
  loc_00E0148D: push esi
  loc_00E0148E: push edi
  loc_00E0148F: mov var_C, esp
  loc_00E01492: mov var_8, 00401C70h
  loc_00E01499: mov var_4, 00000000h
  loc_00E014A0: mov esi, Me
  loc_00E014A3: push esi
  loc_00E014A4: mov eax, [esi]
  loc_00E014A6: call [eax+00000004h]
  loc_00E014A9: mov ecx, New_Enabled
  loc_00E014AC: mov eax, [esi+00000010h]
  loc_00E014AF: mov [esi+0000007Ch], cx
  loc_00E014B3: push ecx
  loc_00E014B4: mov edx, [eax]
  loc_00E014B6: push eax
  loc_00E014B7: call [edx+00000094h]
  loc_00E014BD: test eax, eax
  loc_00E014BF: fnclex
  loc_00E014C1: jge 00E014D8h
  loc_00E014C3: mov ecx, [esi+00000010h]
  loc_00E014C6: push 00000094h
  loc_00E014CB: push 006DCB00h
  loc_00E014D0: push ecx
  loc_00E014D1: push eax
  loc_00E014D2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E014D8: mov edx, [esi]
  loc_00E014DA: push esi
  loc_00E014DB: call [edx+00000910h]
  loc_00E014E1: sub esp, 00000010h
  loc_00E014E4: mov ecx, 00000008h
  loc_00E014E9: mov edi, esp
  loc_00E014EB: mov edx, [esi]
  loc_00E014ED: mov eax, 006DEBBCh ; "Enabled"
  loc_00E014F2: push esi
  loc_00E014F3: mov [edi], ecx
  loc_00E014F5: mov ecx, var_20
  loc_00E014F8: mov [edi+00000004h], ecx
  loc_00E014FB: mov [edi+00000008h], eax
  loc_00E014FE: mov eax, var_18
  loc_00E01501: mov [edi+0000000Ch], eax
  loc_00E01504: call [edx+00000390h]
  loc_00E0150A: test eax, eax
  loc_00E0150C: fnclex
  loc_00E0150E: jge 00E01522h
  loc_00E01510: push 00000390h
  loc_00E01515: push 006DCB00h
  loc_00E0151A: push esi
  loc_00E0151B: push eax
  loc_00E0151C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01522: mov eax, Me
  loc_00E01525: push eax
  loc_00E01526: mov ecx, [eax]
  loc_00E01528: call [ecx+00000008h]
  loc_00E0152B: mov eax, var_4
  loc_00E0152E: mov ecx, var_14
  loc_00E01531: pop edi
  loc_00E01532: pop esi
  loc_00E01533: mov fs:[00000000h], ecx
  loc_00E0153A: pop ebx
  loc_00E0153B: mov esp, ebp
  loc_00E0153D: pop ebp
  loc_00E0153E: retn 0008h
End Sub

Public Property Get Font(arg_C) 'E01550
  loc_00E01550: push ebp
  loc_00E01551: mov ebp, esp
  loc_00E01553: sub esp, 0000000Ch
  loc_00E01556: push 00402836h ; __vbaExceptHandler
  loc_00E0155B: mov eax, fs:[00000000h]
  loc_00E01561: push eax
  loc_00E01562: mov fs:[00000000h], esp
  loc_00E01569: sub esp, 00000018h
  loc_00E0156C: push ebx
  loc_00E0156D: push esi
  loc_00E0156E: push edi
  loc_00E0156F: mov var_C, esp
  loc_00E01572: mov var_8, 00401C78h
  loc_00E01579: xor edi, edi
  loc_00E0157B: mov var_4, edi
  loc_00E0157E: mov esi, Me
  loc_00E01581: push esi
  loc_00E01582: mov eax, [esi]
  loc_00E01584: call [eax+00000004h]
  loc_00E01587: mov ecx, arg_C
  loc_00E0158A: lea eax, var_1C
  loc_00E0158D: push eax
  loc_00E0158E: push esi
  loc_00E0158F: mov [ecx], edi
  loc_00E01591: mov edx, [esi]
  loc_00E01593: mov var_18, edi
  loc_00E01596: mov var_1C, edi
  loc_00E01599: call [edx+00000A2Ch]
  loc_00E0159F: cmp eax, edi
  loc_00E015A1: fnclex
  loc_00E015A3: jge 00E015B7h
  loc_00E015A5: push 00000A2Ch
  loc_00E015AA: push 006DD090h
  loc_00E015AF: push esi
  loc_00E015B0: push eax
  loc_00E015B1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E015B7: mov eax, var_1C
  loc_00E015BA: lea ecx, var_18
  loc_00E015BD: push eax
  loc_00E015BE: push ecx
  loc_00E015BF: mov var_1C, edi
  loc_00E015C2: call [004010ACh] ; __vbaObjSet
  loc_00E015C8: push 00E015E9h
  loc_00E015CD: jmp 00E015E8h
  loc_00E015CF: test var_4, 04h
  loc_00E015D3: jz 00E015DEh
  loc_00E015D5: lea ecx, var_18
  loc_00E015D8: call [00401254h] ; __vbaFreeObj
  loc_00E015DE: lea ecx, var_1C
  loc_00E015E1: call [00401254h] ; __vbaFreeObj
  loc_00E015E7: ret
  loc_00E015E8: ret
  loc_00E015E9: mov eax, Me
  loc_00E015EC: push eax
  loc_00E015ED: mov edx, [eax]
  loc_00E015EF: call [edx+00000008h]
  loc_00E015F2: mov eax, arg_C
  loc_00E015F5: mov ecx, var_18
  loc_00E015F8: mov [eax], ecx
  loc_00E015FA: mov eax, var_4
  loc_00E015FD: mov ecx, var_14
  loc_00E01600: pop edi
  loc_00E01601: pop esi
  loc_00E01602: mov fs:[00000000h], ecx
  loc_00E01609: pop ebx
  loc_00E0160A: mov esp, ebp
  loc_00E0160C: pop ebp
  loc_00E0160D: retn 0008h
End Sub

Public Property Set Font(New_Font) 'E01610
  loc_00E01610: push ebp
  loc_00E01611: mov ebp, esp
  loc_00E01613: sub esp, 0000000Ch
  loc_00E01616: push 00402836h ; __vbaExceptHandler
  loc_00E0161B: mov eax, fs:[00000000h]
  loc_00E01621: push eax
  loc_00E01622: mov fs:[00000000h], esp
  loc_00E01629: sub esp, 00000028h
  loc_00E0162C: push ebx
  loc_00E0162D: push esi
  loc_00E0162E: push edi
  loc_00E0162F: mov var_C, esp
  loc_00E01632: mov var_8, 00401C88h
  loc_00E01639: xor ebx, ebx
  loc_00E0163B: mov var_4, ebx
  loc_00E0163E: mov esi, Me
  loc_00E01641: push esi
  loc_00E01642: mov eax, [esi]
  loc_00E01644: call [eax+00000004h]
  loc_00E01647: mov ecx, New_Font
  loc_00E0164A: mov edi, [004010B4h] ; __vbaObjSetAddref
  loc_00E01650: lea edx, var_18
  loc_00E01653: push ecx
  loc_00E01654: push edx
  loc_00E01655: mov var_18, ebx
  loc_00E01658: mov var_1C, ebx
  loc_00E0165B: call edi
  loc_00E0165D: mov eax, var_18
  loc_00E01660: mov edx, [esi]
  loc_00E01662: lea ecx, var_1C
  loc_00E01665: push eax
  loc_00E01666: push ecx
  loc_00E01667: mov var_3C, edx
  loc_00E0166A: call edi
  loc_00E0166C: mov edx, var_3C
  loc_00E0166F: push eax
  loc_00E01670: push esi
  loc_00E01671: call [edx+00000A34h]
  loc_00E01677: cmp eax, ebx
  loc_00E01679: fnclex
  loc_00E0167B: jge 00E01693h
  loc_00E0167D: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01683: push 00000A34h
  loc_00E01688: push 006DD090h
  loc_00E0168D: push esi
  loc_00E0168E: push eax
  loc_00E0168F: call edi
  loc_00E01691: jmp 00E01699h
  loc_00E01693: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01699: lea ecx, var_1C
  loc_00E0169C: call [00401254h] ; __vbaFreeObj
  loc_00E016A2: mov eax, [esi]
  loc_00E016A4: push esi
  loc_00E016A5: call [eax+00000338h]
  loc_00E016AB: cmp eax, ebx
  loc_00E016AD: fnclex
  loc_00E016AF: jge 00E016BFh
  loc_00E016B1: push 00000338h
  loc_00E016B6: push 006DCB00h
  loc_00E016BB: push esi
  loc_00E016BC: push eax
  loc_00E016BD: call edi
  loc_00E016BF: mov ecx, [esi]
  loc_00E016C1: push esi
  loc_00E016C2: call [ecx+00000910h]
  loc_00E016C8: sub esp, 00000010h
  loc_00E016CB: mov ecx, 00000008h
  loc_00E016D0: mov ebx, esp
  loc_00E016D2: mov edx, [esi]
  loc_00E016D4: mov eax, 006DEBACh ; "Font"
  loc_00E016D9: push esi
  loc_00E016DA: mov [ebx], ecx
  loc_00E016DC: mov ecx, var_28
  loc_00E016DF: mov [ebx+00000004h], ecx
  loc_00E016E2: mov [ebx+00000008h], eax
  loc_00E016E5: mov eax, var_20
  loc_00E016E8: mov [ebx+0000000Ch], eax
  loc_00E016EB: call [edx+00000390h]
  loc_00E016F1: test eax, eax
  loc_00E016F3: fnclex
  loc_00E016F5: jge 00E01705h
  loc_00E016F7: push 00000390h
  loc_00E016FC: push 006DCB00h
  loc_00E01701: push esi
  loc_00E01702: push eax
  loc_00E01703: call edi
  loc_00E01705: mov ecx, [esi]
  loc_00E01707: push 006DCC80h
  loc_00E0170C: push esi
  loc_00E0170D: call [ecx+000009F4h]
  loc_00E01713: test eax, eax
  loc_00E01715: jge 00E01725h
  loc_00E01717: push 000009F4h
  loc_00E0171C: push 006DD090h
  loc_00E01721: push esi
  loc_00E01722: push eax
  loc_00E01723: call edi
  loc_00E01725: push 00E01740h
  loc_00E0172A: jmp 00E01736h
  loc_00E0172C: lea ecx, var_1C
  loc_00E0172F: call [00401254h] ; __vbaFreeObj
  loc_00E01735: ret
  loc_00E01736: lea ecx, var_18
  loc_00E01739: call [00401254h] ; __vbaFreeObj
  loc_00E0173F: ret
  loc_00E01740: mov eax, Me
  loc_00E01743: push eax
  loc_00E01744: mov edx, [eax]
  loc_00E01746: call [edx+00000008h]
  loc_00E01749: mov eax, var_4
  loc_00E0174C: mov ecx, var_14
  loc_00E0174F: pop edi
  loc_00E01750: pop esi
  loc_00E01751: mov fs:[00000000h], ecx
  loc_00E01758: pop ebx
  loc_00E01759: mov esp, ebp
  loc_00E0175B: pop ebp
  loc_00E0175C: retn 0008h
End Sub

Public Property Get ForeColor(arg_C) 'E018E0
  loc_00E018E0: push ebp
  loc_00E018E1: mov ebp, esp
  loc_00E018E3: sub esp, 0000000Ch
  loc_00E018E6: push 00402836h ; __vbaExceptHandler
  loc_00E018EB: mov eax, fs:[00000000h]
  loc_00E018F1: push eax
  loc_00E018F2: mov fs:[00000000h], esp
  loc_00E018F9: sub esp, 0000000Ch
  loc_00E018FC: push ebx
  loc_00E018FD: push esi
  loc_00E018FE: push edi
  loc_00E018FF: mov var_C, esp
  loc_00E01902: mov var_8, 00401CA8h
  loc_00E01909: xor edi, edi
  loc_00E0190B: mov var_4, edi
  loc_00E0190E: mov esi, Me
  loc_00E01911: push esi
  loc_00E01912: mov eax, [esi]
  loc_00E01914: call [eax+00000004h]
  loc_00E01917: mov ecx, [esi+00000090h]
  loc_00E0191D: mov var_18, edi
  loc_00E01920: mov var_18, ecx
  loc_00E01923: mov eax, Me
  loc_00E01926: push eax
  loc_00E01927: mov edx, [eax]
  loc_00E01929: call [edx+00000008h]
  loc_00E0192C: mov eax, arg_C
  loc_00E0192F: mov ecx, var_18
  loc_00E01932: mov [eax], ecx
  loc_00E01934: mov eax, var_4
  loc_00E01937: mov ecx, var_14
  loc_00E0193A: pop edi
  loc_00E0193B: pop esi
  loc_00E0193C: mov fs:[00000000h], ecx
  loc_00E01943: pop ebx
  loc_00E01944: mov esp, ebp
  loc_00E01946: pop ebp
  loc_00E01947: retn 0008h
End Sub

Public Property Let ForeColor(New_ForeColor) 'E01950
  loc_00E01950: push ebp
  loc_00E01951: mov ebp, esp
  loc_00E01953: sub esp, 0000000Ch
  loc_00E01956: push 00402836h ; __vbaExceptHandler
  loc_00E0195B: mov eax, fs:[00000000h]
  loc_00E01961: push eax
  loc_00E01962: mov fs:[00000000h], esp
  loc_00E01969: sub esp, 0000001Ch
  loc_00E0196C: push ebx
  loc_00E0196D: push esi
  loc_00E0196E: push edi
  loc_00E0196F: mov var_C, esp
  loc_00E01972: mov var_8, 00401CB0h
  loc_00E01979: mov var_4, 00000000h
  loc_00E01980: mov esi, Me
  loc_00E01983: push esi
  loc_00E01984: mov eax, [esi]
  loc_00E01986: call [eax+00000004h]
  loc_00E01989: mov ecx, New_ForeColor
  loc_00E0198C: mov eax, [esi+00000010h]
  loc_00E0198F: mov [esi+00000090h], ecx
  loc_00E01995: push ecx
  loc_00E01996: mov edx, [eax]
  loc_00E01998: push eax
  loc_00E01999: call [edx+0000006Ch]
  loc_00E0199C: test eax, eax
  loc_00E0199E: fnclex
  loc_00E019A0: jge 00E019B4h
  loc_00E019A2: mov ecx, [esi+00000010h]
  loc_00E019A5: push 0000006Ch
  loc_00E019A7: push 006DCB00h
  loc_00E019AC: push ecx
  loc_00E019AD: push eax
  loc_00E019AE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E019B4: mov edx, [esi]
  loc_00E019B6: push esi
  loc_00E019B7: call [edx+00000910h]
  loc_00E019BD: sub esp, 00000010h
  loc_00E019C0: mov ecx, 00000008h
  loc_00E019C5: mov edi, esp
  loc_00E019C7: mov edx, [esi]
  loc_00E019C9: mov eax, 006DEE94h ; "ForeColor"
  loc_00E019CE: push esi
  loc_00E019CF: mov [edi], ecx
  loc_00E019D1: mov ecx, var_20
  loc_00E019D4: mov [edi+00000004h], ecx
  loc_00E019D7: mov [edi+00000008h], eax
  loc_00E019DA: mov eax, var_18
  loc_00E019DD: mov [edi+0000000Ch], eax
  loc_00E019E0: call [edx+00000390h]
  loc_00E019E6: test eax, eax
  loc_00E019E8: fnclex
  loc_00E019EA: jge 00E019FEh
  loc_00E019EC: push 00000390h
  loc_00E019F1: push 006DCB00h
  loc_00E019F6: push esi
  loc_00E019F7: push eax
  loc_00E019F8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E019FE: mov eax, Me
  loc_00E01A01: push eax
  loc_00E01A02: mov ecx, [eax]
  loc_00E01A04: call [ecx+00000008h]
  loc_00E01A07: mov eax, var_4
  loc_00E01A0A: mov ecx, var_14
  loc_00E01A0D: pop edi
  loc_00E01A0E: pop esi
  loc_00E01A0F: mov fs:[00000000h], ecx
  loc_00E01A16: pop ebx
  loc_00E01A17: mov esp, ebp
  loc_00E01A19: pop ebp
  loc_00E01A1A: retn 0008h
End Sub

Public Property Get ForeColorHover(arg_C) 'E01A20
  loc_00E01A20: push ebp
  loc_00E01A21: mov ebp, esp
  loc_00E01A23: sub esp, 0000000Ch
  loc_00E01A26: push 00402836h ; __vbaExceptHandler
  loc_00E01A2B: mov eax, fs:[00000000h]
  loc_00E01A31: push eax
  loc_00E01A32: mov fs:[00000000h], esp
  loc_00E01A39: sub esp, 0000000Ch
  loc_00E01A3C: push ebx
  loc_00E01A3D: push esi
  loc_00E01A3E: push edi
  loc_00E01A3F: mov var_C, esp
  loc_00E01A42: mov var_8, 00401CB8h
  loc_00E01A49: xor edi, edi
  loc_00E01A4B: mov var_4, edi
  loc_00E01A4E: mov esi, Me
  loc_00E01A51: push esi
  loc_00E01A52: mov eax, [esi]
  loc_00E01A54: call [eax+00000004h]
  loc_00E01A57: mov ecx, [esi+00000094h]
  loc_00E01A5D: mov var_18, edi
  loc_00E01A60: mov var_18, ecx
  loc_00E01A63: mov eax, Me
  loc_00E01A66: push eax
  loc_00E01A67: mov edx, [eax]
  loc_00E01A69: call [edx+00000008h]
  loc_00E01A6C: mov eax, arg_C
  loc_00E01A6F: mov ecx, var_18
  loc_00E01A72: mov [eax], ecx
  loc_00E01A74: mov eax, var_4
  loc_00E01A77: mov ecx, var_14
  loc_00E01A7A: pop edi
  loc_00E01A7B: pop esi
  loc_00E01A7C: mov fs:[00000000h], ecx
  loc_00E01A83: pop ebx
  loc_00E01A84: mov esp, ebp
  loc_00E01A86: pop ebp
  loc_00E01A87: retn 0008h
End Sub

Public Property Let ForeColorHover(New_ForeColorHover) 'E01A90
  loc_00E01A90: push ebp
  loc_00E01A91: mov ebp, esp
  loc_00E01A93: sub esp, 0000000Ch
  loc_00E01A96: push 00402836h ; __vbaExceptHandler
  loc_00E01A9B: mov eax, fs:[00000000h]
  loc_00E01AA1: push eax
  loc_00E01AA2: mov fs:[00000000h], esp
  loc_00E01AA9: sub esp, 0000001Ch
  loc_00E01AAC: push ebx
  loc_00E01AAD: push esi
  loc_00E01AAE: push edi
  loc_00E01AAF: mov var_C, esp
  loc_00E01AB2: mov var_8, 00401CC0h
  loc_00E01AB9: mov var_4, 00000000h
  loc_00E01AC0: mov esi, Me
  loc_00E01AC3: push esi
  loc_00E01AC4: mov eax, [esi]
  loc_00E01AC6: call [eax+00000004h]
  loc_00E01AC9: mov ecx, New_ForeColorHover
  loc_00E01ACC: mov eax, [esi+00000010h]
  loc_00E01ACF: mov [esi+00000094h], ecx
  loc_00E01AD5: push ecx
  loc_00E01AD6: mov edx, [eax]
  loc_00E01AD8: push eax
  loc_00E01AD9: call [edx+0000006Ch]
  loc_00E01ADC: test eax, eax
  loc_00E01ADE: fnclex
  loc_00E01AE0: jge 00E01AF4h
  loc_00E01AE2: mov ecx, [esi+00000010h]
  loc_00E01AE5: push 0000006Ch
  loc_00E01AE7: push 006DCB00h
  loc_00E01AEC: push ecx
  loc_00E01AED: push eax
  loc_00E01AEE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01AF4: mov edx, [esi]
  loc_00E01AF6: push esi
  loc_00E01AF7: call [edx+00000910h]
  loc_00E01AFD: sub esp, 00000010h
  loc_00E01B00: mov ecx, 00000008h
  loc_00E01B05: mov edi, esp
  loc_00E01B07: mov edx, [esi]
  loc_00E01B09: mov eax, 006DEEACh ; "ForeColorHover"
  loc_00E01B0E: push esi
  loc_00E01B0F: mov [edi], ecx
  loc_00E01B11: mov ecx, var_20
  loc_00E01B14: mov [edi+00000004h], ecx
  loc_00E01B17: mov [edi+00000008h], eax
  loc_00E01B1A: mov eax, var_18
  loc_00E01B1D: mov [edi+0000000Ch], eax
  loc_00E01B20: call [edx+00000390h]
  loc_00E01B26: test eax, eax
  loc_00E01B28: fnclex
  loc_00E01B2A: jge 00E01B3Eh
  loc_00E01B2C: push 00000390h
  loc_00E01B31: push 006DCB00h
  loc_00E01B36: push esi
  loc_00E01B37: push eax
  loc_00E01B38: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01B3E: mov eax, Me
  loc_00E01B41: push eax
  loc_00E01B42: mov ecx, [eax]
  loc_00E01B44: call [ecx+00000008h]
  loc_00E01B47: mov eax, var_4
  loc_00E01B4A: mov ecx, var_14
  loc_00E01B4D: pop edi
  loc_00E01B4E: pop esi
  loc_00E01B4F: mov fs:[00000000h], ecx
  loc_00E01B56: pop ebx
  loc_00E01B57: mov esp, ebp
  loc_00E01B59: pop ebp
  loc_00E01B5A: retn 0008h
End Sub

Public Property Get HandPointer(arg_C) 'E01B60
  loc_00E01B60: push ebp
  loc_00E01B61: mov ebp, esp
  loc_00E01B63: sub esp, 0000000Ch
  loc_00E01B66: push 00402836h ; __vbaExceptHandler
  loc_00E01B6B: mov eax, fs:[00000000h]
  loc_00E01B71: push eax
  loc_00E01B72: mov fs:[00000000h], esp
  loc_00E01B79: sub esp, 0000000Ch
  loc_00E01B7C: push ebx
  loc_00E01B7D: push esi
  loc_00E01B7E: push edi
  loc_00E01B7F: mov var_C, esp
  loc_00E01B82: mov var_8, 00401CC8h
  loc_00E01B89: xor edi, edi
  loc_00E01B8B: mov var_4, edi
  loc_00E01B8E: mov esi, Me
  loc_00E01B91: push esi
  loc_00E01B92: mov eax, [esi]
  loc_00E01B94: call [eax+00000004h]
  loc_00E01B97: mov cx, [esi+00000052h]
  loc_00E01B9B: mov var_18, edi
  loc_00E01B9E: mov var_18, ecx
  loc_00E01BA1: mov eax, Me
  loc_00E01BA4: push eax
  loc_00E01BA5: mov edx, [eax]
  loc_00E01BA7: call [edx+00000008h]
  loc_00E01BAA: mov eax, arg_C
  loc_00E01BAD: mov cx, var_18
  loc_00E01BB1: mov [eax], cx
  loc_00E01BB4: mov eax, var_4
  loc_00E01BB7: mov ecx, var_14
  loc_00E01BBA: pop edi
  loc_00E01BBB: pop esi
  loc_00E01BBC: mov fs:[00000000h], ecx
  loc_00E01BC3: pop ebx
  loc_00E01BC4: mov esp, ebp
  loc_00E01BC6: pop ebp
  loc_00E01BC7: retn 0008h
End Sub

Public Property Let HandPointer(New_HandPointer) 'E01BD0
  loc_00E01BD0: push ebp
  loc_00E01BD1: mov ebp, esp
  loc_00E01BD3: sub esp, 0000000Ch
  loc_00E01BD6: push 00402836h ; __vbaExceptHandler
  loc_00E01BDB: mov eax, fs:[00000000h]
  loc_00E01BE1: push eax
  loc_00E01BE2: mov fs:[00000000h], esp
  loc_00E01BE9: sub esp, 0000001Ch
  loc_00E01BEC: push ebx
  loc_00E01BED: push esi
  loc_00E01BEE: push edi
  loc_00E01BEF: mov var_C, esp
  loc_00E01BF2: mov var_8, 00401CD0h
  loc_00E01BF9: mov var_4, 00000000h
  loc_00E01C00: mov esi, Me
  loc_00E01C03: push esi
  loc_00E01C04: mov eax, [esi]
  loc_00E01C06: call [eax+00000004h]
  loc_00E01C09: mov ax, New_HandPointer
  loc_00E01C0D: test ax, ax
  loc_00E01C10: mov [esi+00000052h], ax
  loc_00E01C14: jz 00E01C3Fh
  loc_00E01C16: mov eax, [esi+00000010h]
  loc_00E01C19: push 00000000h
  loc_00E01C1B: push eax
  loc_00E01C1C: mov ecx, [eax]
  loc_00E01C1E: call [ecx+000000A4h]
  loc_00E01C24: test eax, eax
  loc_00E01C26: fnclex
  loc_00E01C28: jge 00E01C3Fh
  loc_00E01C2A: mov edx, [esi+00000010h]
  loc_00E01C2D: push 000000A4h
  loc_00E01C32: push 006DCB00h
  loc_00E01C37: push edx
  loc_00E01C38: push eax
  loc_00E01C39: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01C3F: mov eax, [esi]
  loc_00E01C41: push esi
  loc_00E01C42: call [eax+00000910h]
  loc_00E01C48: sub esp, 00000010h
  loc_00E01C4B: mov ecx, 00000008h
  loc_00E01C50: mov edi, esp
  loc_00E01C52: mov edx, [esi]
  loc_00E01C54: mov eax, 006DEC2Ch ; "HandPointer"
  loc_00E01C59: push esi
  loc_00E01C5A: mov [edi], ecx
  loc_00E01C5C: mov ecx, var_20
  loc_00E01C5F: mov [edi+00000004h], ecx
  loc_00E01C62: mov [edi+00000008h], eax
  loc_00E01C65: mov eax, var_18
  loc_00E01C68: mov [edi+0000000Ch], eax
  loc_00E01C6B: call [edx+00000390h]
  loc_00E01C71: test eax, eax
  loc_00E01C73: fnclex
  loc_00E01C75: jge 00E01C89h
  loc_00E01C77: push 00000390h
  loc_00E01C7C: push 006DCB00h
  loc_00E01C81: push esi
  loc_00E01C82: push eax
  loc_00E01C83: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01C89: mov eax, Me
  loc_00E01C8C: push eax
  loc_00E01C8D: mov ecx, [eax]
  loc_00E01C8F: call [ecx+00000008h]
  loc_00E01C92: mov eax, var_4
  loc_00E01C95: mov ecx, var_14
  loc_00E01C98: pop edi
  loc_00E01C99: pop esi
  loc_00E01C9A: mov fs:[00000000h], ecx
  loc_00E01CA1: pop ebx
  loc_00E01CA2: mov esp, ebp
  loc_00E01CA4: pop ebp
  loc_00E01CA5: retn 0008h
End Sub

Public Property Get Hwnd(arg_C) 'E01CB0
  loc_00E01CB0: push ebp
  loc_00E01CB1: mov ebp, esp
  loc_00E01CB3: sub esp, 0000000Ch
  loc_00E01CB6: push 00402836h ; __vbaExceptHandler
  loc_00E01CBB: mov eax, fs:[00000000h]
  loc_00E01CC1: push eax
  loc_00E01CC2: mov fs:[00000000h], esp
  loc_00E01CC9: sub esp, 00000014h
  loc_00E01CCC: push ebx
  loc_00E01CCD: push esi
  loc_00E01CCE: push edi
  loc_00E01CCF: mov var_C, esp
  loc_00E01CD2: mov var_8, 00401CD8h
  loc_00E01CD9: xor edi, edi
  loc_00E01CDB: mov var_4, edi
  loc_00E01CDE: mov esi, Me
  loc_00E01CE1: push esi
  loc_00E01CE2: mov eax, [esi]
  loc_00E01CE4: call [eax+00000004h]
  loc_00E01CE7: mov eax, [esi+00000010h]
  loc_00E01CEA: lea edx, var_1C
  loc_00E01CED: mov var_1C, edi
  loc_00E01CF0: push edx
  loc_00E01CF1: mov ecx, [eax]
  loc_00E01CF3: push eax
  loc_00E01CF4: mov var_18, edi
  loc_00E01CF7: call [ecx+00000058h]
  loc_00E01CFA: cmp eax, edi
  loc_00E01CFC: fnclex
  loc_00E01CFE: jge 00E01D12h
  loc_00E01D00: mov ecx, [esi+00000010h]
  loc_00E01D03: push 00000058h
  loc_00E01D05: push 006DCB00h
  loc_00E01D0A: push ecx
  loc_00E01D0B: push eax
  loc_00E01D0C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01D12: mov edx, var_1C
  loc_00E01D15: mov var_18, edx
  loc_00E01D18: mov eax, Me
  loc_00E01D1B: push eax
  loc_00E01D1C: mov ecx, [eax]
  loc_00E01D1E: call [ecx+00000008h]
  loc_00E01D21: mov edx, arg_C
  loc_00E01D24: mov eax, var_18
  loc_00E01D27: mov [edx], eax
  loc_00E01D29: mov eax, var_4
  loc_00E01D2C: mov ecx, var_14
  loc_00E01D2F: pop edi
  loc_00E01D30: pop esi
  loc_00E01D31: mov fs:[00000000h], ecx
  loc_00E01D38: pop ebx
  loc_00E01D39: mov esp, ebp
  loc_00E01D3B: pop ebp
  loc_00E01D3C: retn 0008h
End Sub

Public Property Get MaskColor(arg_C) 'E01D40
  loc_00E01D40: push ebp
  loc_00E01D41: mov ebp, esp
  loc_00E01D43: sub esp, 0000000Ch
  loc_00E01D46: push 00402836h ; __vbaExceptHandler
  loc_00E01D4B: mov eax, fs:[00000000h]
  loc_00E01D51: push eax
  loc_00E01D52: mov fs:[00000000h], esp
  loc_00E01D59: sub esp, 0000000Ch
  loc_00E01D5C: push ebx
  loc_00E01D5D: push esi
  loc_00E01D5E: push edi
  loc_00E01D5F: mov var_C, esp
  loc_00E01D62: mov var_8, 00401CE0h
  loc_00E01D69: xor edi, edi
  loc_00E01D6B: mov var_4, edi
  loc_00E01D6E: mov esi, Me
  loc_00E01D71: push esi
  loc_00E01D72: mov eax, [esi]
  loc_00E01D74: call [eax+00000004h]
  loc_00E01D77: mov ecx, [esi+000000A0h]
  loc_00E01D7D: mov var_18, edi
  loc_00E01D80: mov var_18, ecx
  loc_00E01D83: mov eax, Me
  loc_00E01D86: push eax
  loc_00E01D87: mov edx, [eax]
  loc_00E01D89: call [edx+00000008h]
  loc_00E01D8C: mov eax, arg_C
  loc_00E01D8F: mov ecx, var_18
  loc_00E01D92: mov [eax], ecx
  loc_00E01D94: mov eax, var_4
  loc_00E01D97: mov ecx, var_14
  loc_00E01D9A: pop edi
  loc_00E01D9B: pop esi
  loc_00E01D9C: mov fs:[00000000h], ecx
  loc_00E01DA3: pop ebx
  loc_00E01DA4: mov esp, ebp
  loc_00E01DA6: pop ebp
  loc_00E01DA7: retn 0008h
End Sub

Public Property Let MaskColor(New_MaskColor) 'E01DB0
  loc_00E01DB0: push ebp
  loc_00E01DB1: mov ebp, esp
  loc_00E01DB3: sub esp, 0000000Ch
  loc_00E01DB6: push 00402836h ; __vbaExceptHandler
  loc_00E01DBB: mov eax, fs:[00000000h]
  loc_00E01DC1: push eax
  loc_00E01DC2: mov fs:[00000000h], esp
  loc_00E01DC9: sub esp, 0000001Ch
  loc_00E01DCC: push ebx
  loc_00E01DCD: push esi
  loc_00E01DCE: push edi
  loc_00E01DCF: mov var_C, esp
  loc_00E01DD2: mov var_8, 00401CE8h
  loc_00E01DD9: mov var_4, 00000000h
  loc_00E01DE0: mov esi, Me
  loc_00E01DE3: push esi
  loc_00E01DE4: mov eax, [esi]
  loc_00E01DE6: call [eax+00000004h]
  loc_00E01DE9: mov ecx, New_MaskColor
  loc_00E01DEC: mov edx, [esi]
  loc_00E01DEE: push esi
  loc_00E01DEF: mov [esi+000000A0h], ecx
  loc_00E01DF5: call [edx+00000910h]
  loc_00E01DFB: sub esp, 00000010h
  loc_00E01DFE: mov ecx, 00000008h
  loc_00E01E03: mov edi, esp
  loc_00E01E05: mov edx, [esi]
  loc_00E01E07: mov eax, 006DEDE8h ; "MaskColor"
  loc_00E01E0C: push esi
  loc_00E01E0D: mov [edi], ecx
  loc_00E01E0F: mov ecx, var_20
  loc_00E01E12: mov [edi+00000004h], ecx
  loc_00E01E15: mov [edi+00000008h], eax
  loc_00E01E18: mov eax, var_18
  loc_00E01E1B: mov [edi+0000000Ch], eax
  loc_00E01E1E: call [edx+00000390h]
  loc_00E01E24: test eax, eax
  loc_00E01E26: fnclex
  loc_00E01E28: jge 00E01E3Ch
  loc_00E01E2A: push 00000390h
  loc_00E01E2F: push 006DCB00h
  loc_00E01E34: push esi
  loc_00E01E35: push eax
  loc_00E01E36: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01E3C: mov eax, Me
  loc_00E01E3F: push eax
  loc_00E01E40: mov ecx, [eax]
  loc_00E01E42: call [ecx+00000008h]
  loc_00E01E45: mov eax, var_4
  loc_00E01E48: mov ecx, var_14
  loc_00E01E4B: pop edi
  loc_00E01E4C: pop esi
  loc_00E01E4D: mov fs:[00000000h], ecx
  loc_00E01E54: pop ebx
  loc_00E01E55: mov esp, ebp
  loc_00E01E57: pop ebp
  loc_00E01E58: retn 0008h
End Sub

Public Property Get Mode(arg_C) 'E01E60
  loc_00E01E60: push ebp
  loc_00E01E61: mov ebp, esp
  loc_00E01E63: sub esp, 0000000Ch
  loc_00E01E66: push 00402836h ; __vbaExceptHandler
  loc_00E01E6B: mov eax, fs:[00000000h]
  loc_00E01E71: push eax
  loc_00E01E72: mov fs:[00000000h], esp
  loc_00E01E79: sub esp, 0000000Ch
  loc_00E01E7C: push ebx
  loc_00E01E7D: push esi
  loc_00E01E7E: push edi
  loc_00E01E7F: mov var_C, esp
  loc_00E01E82: mov var_8, 00401CF0h
  loc_00E01E89: xor edi, edi
  loc_00E01E8B: mov var_4, edi
  loc_00E01E8E: mov esi, Me
  loc_00E01E91: push esi
  loc_00E01E92: mov eax, [esi]
  loc_00E01E94: call [eax+00000004h]
  loc_00E01E97: mov ecx, [esi+00000064h]
  loc_00E01E9A: mov var_18, edi
  loc_00E01E9D: mov var_18, ecx
  loc_00E01EA0: mov eax, Me
  loc_00E01EA3: push eax
  loc_00E01EA4: mov edx, [eax]
  loc_00E01EA6: call [edx+00000008h]
  loc_00E01EA9: mov eax, arg_C
  loc_00E01EAC: mov ecx, var_18
  loc_00E01EAF: mov [eax], ecx
  loc_00E01EB1: mov eax, var_4
  loc_00E01EB4: mov ecx, var_14
  loc_00E01EB7: pop edi
  loc_00E01EB8: pop esi
  loc_00E01EB9: mov fs:[00000000h], ecx
  loc_00E01EC0: pop ebx
  loc_00E01EC1: mov esp, ebp
  loc_00E01EC3: pop ebp
  loc_00E01EC4: retn 0008h
End Sub

Public Property Let Mode(New_mode) 'E01ED0
  loc_00E01ED0: push ebp
  loc_00E01ED1: mov ebp, esp
  loc_00E01ED3: sub esp, 0000000Ch
  loc_00E01ED6: push 00402836h ; __vbaExceptHandler
  loc_00E01EDB: mov eax, fs:[00000000h]
  loc_00E01EE1: push eax
  loc_00E01EE2: mov fs:[00000000h], esp
  loc_00E01EE9: sub esp, 00000020h
  loc_00E01EEC: push ebx
  loc_00E01EED: push esi
  loc_00E01EEE: push edi
  loc_00E01EEF: mov var_C, esp
  loc_00E01EF2: mov var_8, 00401CF8h
  loc_00E01EF9: xor ebx, ebx
  loc_00E01EFB: mov var_4, ebx
  loc_00E01EFE: mov esi, Me
  loc_00E01F01: push esi
  loc_00E01F02: mov eax, [esi]
  loc_00E01F04: call [eax+00000004h]
  loc_00E01F07: mov eax, New_mode
  loc_00E01F0A: cmp eax, ebx
  loc_00E01F0C: mov [esi+00000064h], eax
  loc_00E01F0F: jnz 00E01F14h
  loc_00E01F11: mov [esi+00000048h], ebx
  loc_00E01F14: mov ecx, [esi]
  loc_00E01F16: push esi
  loc_00E01F17: call [ecx+00000910h]
  loc_00E01F1D: mov edx, [esi]
  loc_00E01F1F: mov edi, var_20
  loc_00E01F22: sub esp, 00000010h
  loc_00E01F25: mov var_34, edx
  loc_00E01F28: mov edx, esp
  loc_00E01F2A: mov ecx, 00000008h
  loc_00E01F2F: mov eax, 006DEBFCh ; "Value"
  loc_00E01F34: push esi
  loc_00E01F35: mov [edx], ecx
  loc_00E01F37: mov ecx, var_34
  loc_00E01F3A: mov [edx+00000004h], edi
  loc_00E01F3D: mov [edx+00000008h], eax
  loc_00E01F40: mov eax, var_18
  loc_00E01F43: mov [edx+0000000Ch], eax
  loc_00E01F46: call [ecx+00000390h]
  loc_00E01F4C: cmp eax, ebx
  loc_00E01F4E: fnclex
  loc_00E01F50: jge 00E01F64h
  loc_00E01F52: push 00000390h
  loc_00E01F57: push 006DCB00h
  loc_00E01F5C: push esi
  loc_00E01F5D: push eax
  loc_00E01F5E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01F64: sub esp, 00000010h
  loc_00E01F67: mov ecx, 00000008h
  loc_00E01F6C: mov ebx, esp
  loc_00E01F6E: mov edx, [esi]
  loc_00E01F70: mov eax, 006DEE44h ; "Mode"
  loc_00E01F75: push esi
  loc_00E01F76: mov [ebx], ecx
  loc_00E01F78: mov [ebx+00000004h], edi
  loc_00E01F7B: mov [ebx+00000008h], eax
  loc_00E01F7E: mov eax, var_18
  loc_00E01F81: mov [ebx+0000000Ch], eax
  loc_00E01F84: call [edx+00000390h]
  loc_00E01F8A: test eax, eax
  loc_00E01F8C: fnclex
  loc_00E01F8E: jge 00E01FA2h
  loc_00E01F90: push 00000390h
  loc_00E01F95: push 006DCB00h
  loc_00E01F9A: push esi
  loc_00E01F9B: push eax
  loc_00E01F9C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E01FA2: mov eax, Me
  loc_00E01FA5: push eax
  loc_00E01FA6: mov ecx, [eax]
  loc_00E01FA8: call [ecx+00000008h]
  loc_00E01FAB: mov eax, var_4
  loc_00E01FAE: mov ecx, var_14
  loc_00E01FB1: pop edi
  loc_00E01FB2: pop esi
  loc_00E01FB3: mov fs:[00000000h], ecx
  loc_00E01FBA: pop ebx
  loc_00E01FBB: mov esp, ebp
  loc_00E01FBD: pop ebp
  loc_00E01FBE: retn 0008h
End Sub

Public Property Get MouseIcon(arg_C) 'E01FD0
  loc_00E01FD0: push ebp
  loc_00E01FD1: mov ebp, esp
  loc_00E01FD3: sub esp, 0000000Ch
  loc_00E01FD6: push 00402836h ; __vbaExceptHandler
  loc_00E01FDB: mov eax, fs:[00000000h]
  loc_00E01FE1: push eax
  loc_00E01FE2: mov fs:[00000000h], esp
  loc_00E01FE9: sub esp, 00000018h
  loc_00E01FEC: push ebx
  loc_00E01FED: push esi
  loc_00E01FEE: push edi
  loc_00E01FEF: mov var_C, esp
  loc_00E01FF2: mov var_8, 00401D00h
  loc_00E01FF9: xor edi, edi
  loc_00E01FFB: mov var_4, edi
  loc_00E01FFE: mov esi, Me
  loc_00E02001: push esi
  loc_00E02002: mov eax, [esi]
  loc_00E02004: call [eax+00000004h]
  loc_00E02007: mov ecx, arg_C
  loc_00E0200A: mov var_18, edi
  loc_00E0200D: mov var_1C, edi
  loc_00E02010: mov [ecx], edi
  loc_00E02012: mov eax, [esi+00000010h]
  loc_00E02015: lea ecx, var_1C
  loc_00E02018: mov edx, [eax]
  loc_00E0201A: push ecx
  loc_00E0201B: push eax
  loc_00E0201C: call [edx+00000220h]
  loc_00E02022: cmp eax, edi
  loc_00E02024: fnclex
  loc_00E02026: jge 00E0203Dh
  loc_00E02028: mov edx, [esi+00000010h]
  loc_00E0202B: push 00000220h
  loc_00E02030: push 006DCB00h
  loc_00E02035: push edx
  loc_00E02036: push eax
  loc_00E02037: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0203D: mov eax, var_1C
  loc_00E02040: mov var_1C, edi
  loc_00E02043: push eax
  loc_00E02044: lea eax, var_18
  loc_00E02047: push eax
  loc_00E02048: call [004010ACh] ; __vbaObjSet
  loc_00E0204E: push 00E0206Fh
  loc_00E02053: jmp 00E0206Eh
  loc_00E02055: test var_4, 04h
  loc_00E02059: jz 00E02064h
  loc_00E0205B: lea ecx, var_18
  loc_00E0205E: call [00401254h] ; __vbaFreeObj
  loc_00E02064: lea ecx, var_1C
  loc_00E02067: call [00401254h] ; __vbaFreeObj
  loc_00E0206D: ret
  loc_00E0206E: ret
  loc_00E0206F: mov eax, Me
  loc_00E02072: push eax
  loc_00E02073: mov ecx, [eax]
  loc_00E02075: call [ecx+00000008h]
  loc_00E02078: mov edx, arg_C
  loc_00E0207B: mov eax, var_18
  loc_00E0207E: mov [edx], eax
  loc_00E02080: mov eax, var_4
  loc_00E02083: mov ecx, var_14
  loc_00E02086: pop edi
  loc_00E02087: pop esi
  loc_00E02088: mov fs:[00000000h], ecx
  loc_00E0208F: pop ebx
  loc_00E02090: mov esp, ebp
  loc_00E02092: pop ebp
  loc_00E02093: retn 0008h
End Sub

Public Property Set MouseIcon(New_Icon) 'E020A0
  loc_00E020A0: push ebp
  loc_00E020A1: mov ebp, esp
  loc_00E020A3: sub esp, 00000018h
  loc_00E020A6: push 00402836h ; __vbaExceptHandler
  loc_00E020AB: mov eax, fs:[00000000h]
  loc_00E020B1: push eax
  loc_00E020B2: mov fs:[00000000h], esp
  loc_00E020B9: mov eax, 00000044h
  loc_00E020BE: call 00402830h ; __vbaChkstk
  loc_00E020C3: push ebx
  loc_00E020C4: push esi
  loc_00E020C5: push edi
  loc_00E020C6: mov var_18, esp
  loc_00E020C9: mov var_14, 00401D10h ; "'"
  loc_00E020D0: mov var_10, 00000000h
  loc_00E020D7: mov var_C, 00000000h
  loc_00E020DE: mov eax, Me
  loc_00E020E1: mov ecx, [eax]
  loc_00E020E3: mov edx, Me
  loc_00E020E6: push edx
  loc_00E020E7: call [ecx+00000004h]
  loc_00E020EA: mov var_4, 00000001h
  loc_00E020F1: mov eax, New_Icon
  loc_00E020F4: push eax
  loc_00E020F5: lea ecx, var_24
  loc_00E020F8: push ecx
  loc_00E020F9: call [004010B4h] ; __vbaObjSetAddref
  loc_00E020FF: mov var_4, 00000002h
  loc_00E02106: push FFFFFFFFh
  loc_00E02108: call [004010A4h] ; __vbaOnError
  loc_00E0210E: mov var_4, 00000003h
  loc_00E02115: mov edx, var_24
  loc_00E02118: push edx
  loc_00E02119: lea eax, var_28
  loc_00E0211C: push eax
  loc_00E0211D: call [004010B4h] ; __vbaObjSetAddref
  loc_00E02123: push eax
  loc_00E02124: mov ecx, Me
  loc_00E02127: mov edx, [ecx+00000010h]
  loc_00E0212A: mov eax, Me
  loc_00E0212D: mov ecx, [eax+00000010h]
  loc_00E02130: mov eax, [ecx]
  loc_00E02132: push edx
  loc_00E02133: call [eax+00000224h]
  loc_00E02139: fnclex
  loc_00E0213B: mov var_3C, eax
  loc_00E0213E: cmp var_3C, 00000000h
  loc_00E02142: jge 00E02164h
  loc_00E02144: push 00000224h
  loc_00E02149: push 006DCB00h
  loc_00E0214E: mov ecx, Me
  loc_00E02151: mov edx, [ecx+00000010h]
  loc_00E02154: push edx
  loc_00E02155: mov eax, var_3C
  loc_00E02158: push eax
  loc_00E02159: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0215F: mov var_54, eax
  loc_00E02162: jmp 00E0216Bh
  loc_00E02164: mov var_54, 00000000h
  loc_00E0216B: lea ecx, var_28
  loc_00E0216E: call [00401254h] ; __vbaFreeObj
  loc_00E02174: mov var_4, 00000004h
  loc_00E0217B: push 00000000h
  loc_00E0217D: mov ecx, var_24
  loc_00E02180: push ecx
  loc_00E02181: call [0040114Ch] ; __vbaObjIs
  loc_00E02187: movsx edx, ax
  loc_00E0218A: test edx, edx
  loc_00E0218C: jz 00E021E3h
  loc_00E0218E: mov var_4, 00000005h
  loc_00E02195: push 00000000h
  loc_00E02197: mov eax, Me
  loc_00E0219A: mov ecx, [eax+00000010h]
  loc_00E0219D: mov edx, Me
  loc_00E021A0: mov eax, [edx+00000010h]
  loc_00E021A3: mov edx, [eax]
  loc_00E021A5: push ecx
  loc_00E021A6: call [edx+000000A4h]
  loc_00E021AC: fnclex
  loc_00E021AE: mov var_3C, eax
  loc_00E021B1: cmp var_3C, 00000000h
  loc_00E021B5: jge 00E021D7h
  loc_00E021B7: push 000000A4h
  loc_00E021BC: push 006DCB00h
  loc_00E021C1: mov eax, Me
  loc_00E021C4: mov ecx, [eax+00000010h]
  loc_00E021C7: push ecx
  loc_00E021C8: mov edx, var_3C
  loc_00E021CB: push edx
  loc_00E021CC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E021D2: mov var_58, eax
  loc_00E021D5: jmp 00E021DEh
  loc_00E021D7: mov var_58, 00000000h
  loc_00E021DE: jmp 00E022B9h
  loc_00E021E3: mov var_4, 00000007h
  loc_00E021EA: mov eax, Me
  loc_00E021ED: mov [eax+00000052h], 0000h
  loc_00E021F3: mov var_4, 00000008h
  loc_00E021FA: mov var_30, 006DEC2Ch ; "HandPointer"
  loc_00E02201: mov var_38, 00000008h
  loc_00E02208: mov eax, 00000010h
  loc_00E0220D: call 00402830h ; __vbaChkstk
  loc_00E02212: mov ecx, esp
  loc_00E02214: mov edx, var_38
  loc_00E02217: mov [ecx], edx
  loc_00E02219: mov eax, var_34
  loc_00E0221C: mov [ecx+00000004h], eax
  loc_00E0221F: mov edx, var_30
  loc_00E02222: mov [ecx+00000008h], edx
  loc_00E02225: mov eax, var_2C
  loc_00E02228: mov [ecx+0000000Ch], eax
  loc_00E0222B: mov ecx, Me
  loc_00E0222E: mov edx, [ecx]
  loc_00E02230: mov eax, Me
  loc_00E02233: push eax
  loc_00E02234: call [edx+00000390h]
  loc_00E0223A: fnclex
  loc_00E0223C: mov var_3C, eax
  loc_00E0223F: cmp var_3C, 00000000h
  loc_00E02243: jge 00E02262h
  loc_00E02245: push 00000390h
  loc_00E0224A: push 006DCB00h
  loc_00E0224F: mov ecx, Me
  loc_00E02252: push ecx
  loc_00E02253: mov edx, var_3C
  loc_00E02256: push edx
  loc_00E02257: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0225D: mov var_5C, eax
  loc_00E02260: jmp 00E02269h
  loc_00E02262: mov var_5C, 00000000h
  loc_00E02269: mov var_4, 00000009h
  loc_00E02270: push 00000063h
  loc_00E02272: mov eax, Me
  loc_00E02275: mov ecx, [eax+00000010h]
  loc_00E02278: mov edx, Me
  loc_00E0227B: mov eax, [edx+00000010h]
  loc_00E0227E: mov edx, [eax]
  loc_00E02280: push ecx
  loc_00E02281: call [edx+000000A4h]
  loc_00E02287: fnclex
  loc_00E02289: mov var_3C, eax
  loc_00E0228C: cmp var_3C, 00000000h
  loc_00E02290: jge 00E022B2h
  loc_00E02292: push 000000A4h
  loc_00E02297: push 006DCB00h
  loc_00E0229C: mov eax, Me
  loc_00E0229F: mov ecx, [eax+00000010h]
  loc_00E022A2: push ecx
  loc_00E022A3: mov edx, var_3C
  loc_00E022A6: push edx
  loc_00E022A7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E022AD: mov var_60, eax
  loc_00E022B0: jmp 00E022B9h
  loc_00E022B2: mov var_60, 00000000h
  loc_00E022B9: mov var_4, 0000000Bh
  loc_00E022C0: mov var_30, 006DEC54h ; "MouseIcon"
  loc_00E022C7: mov var_38, 00000008h
  loc_00E022CE: mov eax, 00000010h
  loc_00E022D3: call 00402830h ; __vbaChkstk
  loc_00E022D8: mov eax, esp
  loc_00E022DA: mov ecx, var_38
  loc_00E022DD: mov [eax], ecx
  loc_00E022DF: mov edx, var_34
  loc_00E022E2: mov [eax+00000004h], edx
  loc_00E022E5: mov ecx, var_30
  loc_00E022E8: mov [eax+00000008h], ecx
  loc_00E022EB: mov edx, var_2C
  loc_00E022EE: mov [eax+0000000Ch], edx
  loc_00E022F1: mov eax, Me
  loc_00E022F4: mov ecx, [eax]
  loc_00E022F6: mov edx, Me
  loc_00E022F9: push edx
  loc_00E022FA: call [ecx+00000390h]
  loc_00E02300: fnclex
  loc_00E02302: mov var_3C, eax
  loc_00E02305: cmp var_3C, 00000000h
  loc_00E02309: jge 00E02328h
  loc_00E0230B: push 00000390h
  loc_00E02310: push 006DCB00h
  loc_00E02315: mov eax, Me
  loc_00E02318: push eax
  loc_00E02319: mov ecx, var_3C
  loc_00E0231C: push ecx
  loc_00E0231D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E02323: mov var_64, eax
  loc_00E02326: jmp 00E0232Fh
  loc_00E02328: mov var_64, 00000000h
  loc_00E0232F: push 00E0234Ah
  loc_00E02334: jmp 00E02340h
  loc_00E02336: lea ecx, var_28
  loc_00E02339: call [00401254h] ; __vbaFreeObj
  loc_00E0233F: ret
  loc_00E02340: lea ecx, var_24
  loc_00E02343: call [00401254h] ; __vbaFreeObj
  loc_00E02349: ret
  loc_00E0234A: mov edx, Me
  loc_00E0234D: mov eax, [edx]
  loc_00E0234F: mov ecx, Me
  loc_00E02352: push ecx
  loc_00E02353: call [eax+00000008h]
  loc_00E02356: mov eax, var_10
  loc_00E02359: mov ecx, var_20
  loc_00E0235C: mov fs:[00000000h], ecx
  loc_00E02363: pop edi
  loc_00E02364: pop esi
  loc_00E02365: pop ebx
  loc_00E02366: mov esp, ebp
  loc_00E02368: pop ebp
  loc_00E02369: retn 0008h
End Sub

Public Property Get MousePointer(arg_C) 'E02370
  loc_00E02370: push ebp
  loc_00E02371: mov ebp, esp
  loc_00E02373: sub esp, 0000000Ch
  loc_00E02376: push 00402836h ; __vbaExceptHandler
  loc_00E0237B: mov eax, fs:[00000000h]
  loc_00E02381: push eax
  loc_00E02382: mov fs:[00000000h], esp
  loc_00E02389: sub esp, 00000014h
  loc_00E0238C: push ebx
  loc_00E0238D: push esi
  loc_00E0238E: push edi
  loc_00E0238F: mov var_C, esp
  loc_00E02392: mov var_8, 00401D60h
  loc_00E02399: xor edi, edi
  loc_00E0239B: mov var_4, edi
  loc_00E0239E: mov esi, Me
  loc_00E023A1: push esi
  loc_00E023A2: mov eax, [esi]
  loc_00E023A4: call [eax+00000004h]
  loc_00E023A7: mov eax, [esi+00000010h]
  loc_00E023AA: lea edx, var_1C
  loc_00E023AD: mov var_1C, edi
  loc_00E023B0: push edx
  loc_00E023B1: mov ecx, [eax]
  loc_00E023B3: push eax
  loc_00E023B4: mov var_18, edi
  loc_00E023B7: call [ecx+000000A0h]
  loc_00E023BD: cmp eax, edi
  loc_00E023BF: fnclex
  loc_00E023C1: jge 00E023D8h
  loc_00E023C3: mov ecx, [esi+00000010h]
  loc_00E023C6: push 000000A0h
  loc_00E023CB: push 006DCB00h
  loc_00E023D0: push ecx
  loc_00E023D1: push eax
  loc_00E023D2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E023D8: movsx edx, var_1C
  loc_00E023DC: mov var_18, edx
  loc_00E023DF: mov eax, Me
  loc_00E023E2: push eax
  loc_00E023E3: mov ecx, [eax]
  loc_00E023E5: call [ecx+00000008h]
  loc_00E023E8: mov edx, arg_C
  loc_00E023EB: mov eax, var_18
  loc_00E023EE: mov [edx], eax
  loc_00E023F0: mov eax, var_4
  loc_00E023F3: mov ecx, var_14
  loc_00E023F6: pop edi
  loc_00E023F7: pop esi
  loc_00E023F8: mov fs:[00000000h], ecx
  loc_00E023FF: pop ebx
  loc_00E02400: mov esp, ebp
  loc_00E02402: pop ebp
  loc_00E02403: retn 0008h
End Sub

Public Property Let MousePointer(New_Cursor) 'E02410
  loc_00E02410: push ebp
  loc_00E02411: mov ebp, esp
  loc_00E02413: sub esp, 0000000Ch
  loc_00E02416: push 00402836h ; __vbaExceptHandler
  loc_00E0241B: mov eax, fs:[00000000h]
  loc_00E02421: push eax
  loc_00E02422: mov fs:[00000000h], esp
  loc_00E02429: sub esp, 0000001Ch
  loc_00E0242C: push ebx
  loc_00E0242D: push esi
  loc_00E0242E: push edi
  loc_00E0242F: mov var_C, esp
  loc_00E02432: mov var_8, 00401D68h
  loc_00E02439: mov var_4, 00000000h
  loc_00E02440: mov esi, Me
  loc_00E02443: push esi
  loc_00E02444: mov eax, [esi]
  loc_00E02446: call [eax+00000004h]
  loc_00E02449: mov ecx, [esi+00000010h]
  loc_00E0244C: mov edi, [ecx]
  loc_00E0244E: mov ecx, New_Cursor
  loc_00E02451: call [00401110h] ; __vbaI2I4
  loc_00E02457: mov edx, [esi+00000010h]
  loc_00E0245A: push eax
  loc_00E0245B: push edx
  loc_00E0245C: call [edi+000000A4h]
  loc_00E02462: test eax, eax
  loc_00E02464: fnclex
  loc_00E02466: jge 00E0247Dh
  loc_00E02468: mov ecx, [esi+00000010h]
  loc_00E0246B: push 000000A4h
  loc_00E02470: push 006DCB00h
  loc_00E02475: push ecx
  loc_00E02476: push eax
  loc_00E02477: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0247D: sub esp, 00000010h
  loc_00E02480: mov ecx, 00000008h
  loc_00E02485: mov edi, esp
  loc_00E02487: mov edx, [esi]
  loc_00E02489: mov eax, 006DEC0Ch ; "MousePointer"
  loc_00E0248E: push esi
  loc_00E0248F: mov [edi], ecx
  loc_00E02491: mov ecx, var_20
  loc_00E02494: mov [edi+00000004h], ecx
  loc_00E02497: mov [edi+00000008h], eax
  loc_00E0249A: mov eax, var_18
  loc_00E0249D: mov [edi+0000000Ch], eax
  loc_00E024A0: call [edx+00000390h]
  loc_00E024A6: test eax, eax
  loc_00E024A8: fnclex
  loc_00E024AA: jge 00E024BEh
  loc_00E024AC: push 00000390h
  loc_00E024B1: push 006DCB00h
  loc_00E024B6: push esi
  loc_00E024B7: push eax
  loc_00E024B8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E024BE: mov eax, Me
  loc_00E024C1: push eax
  loc_00E024C2: mov ecx, [eax]
  loc_00E024C4: call [ecx+00000008h]
  loc_00E024C7: mov eax, var_4
  loc_00E024CA: mov ecx, var_14
  loc_00E024CD: pop edi
  loc_00E024CE: pop esi
  loc_00E024CF: mov fs:[00000000h], ecx
  loc_00E024D6: pop ebx
  loc_00E024D7: mov esp, ebp
  loc_00E024D9: pop ebp
  loc_00E024DA: retn 0008h
End Sub

Public Property Get PictureNormal(arg_C) 'E024E0
  loc_00E024E0: push ebp
  loc_00E024E1: mov ebp, esp
  loc_00E024E3: sub esp, 0000000Ch
  loc_00E024E6: push 00402836h ; __vbaExceptHandler
  loc_00E024EB: mov eax, fs:[00000000h]
  loc_00E024F1: push eax
  loc_00E024F2: mov fs:[00000000h], esp
  loc_00E024F9: sub esp, 0000000Ch
  loc_00E024FC: push ebx
  loc_00E024FD: push esi
  loc_00E024FE: push edi
  loc_00E024FF: mov var_C, esp
  loc_00E02502: mov var_8, 00401D70h
  loc_00E02509: xor edi, edi
  loc_00E0250B: mov var_4, edi
  loc_00E0250E: mov esi, Me
  loc_00E02511: push esi
  loc_00E02512: mov eax, [esi]
  loc_00E02514: call [eax+00000004h]
  loc_00E02517: mov ecx, arg_C
  loc_00E0251A: lea eax, var_18
  loc_00E0251D: mov var_18, edi
  loc_00E02520: mov [ecx], edi
  loc_00E02522: mov edx, [esi+0000014Ch]
  loc_00E02528: push edx
  loc_00E02529: push eax
  loc_00E0252A: call [004010B4h] ; __vbaObjSetAddref
  loc_00E02530: push 00E02542h
  loc_00E02535: jmp 00E02541h
  loc_00E02537: lea ecx, var_18
  loc_00E0253A: call [00401254h] ; __vbaFreeObj
  loc_00E02540: ret
  loc_00E02541: ret
  loc_00E02542: mov eax, Me
  loc_00E02545: push eax
  loc_00E02546: mov ecx, [eax]
  loc_00E02548: call [ecx+00000008h]
  loc_00E0254B: mov edx, arg_C
  loc_00E0254E: mov eax, var_18
  loc_00E02551: mov [edx], eax
  loc_00E02553: mov eax, var_4
  loc_00E02556: mov ecx, var_14
  loc_00E02559: pop edi
  loc_00E0255A: pop esi
  loc_00E0255B: mov fs:[00000000h], ecx
  loc_00E02562: pop ebx
  loc_00E02563: mov esp, ebp
  loc_00E02565: pop ebp
  loc_00E02566: retn 0008h
End Sub

Public Property Set PictureNormal(New_Picture) 'E02570
  loc_00E02570: push ebp
  loc_00E02571: mov ebp, esp
  loc_00E02573: sub esp, 0000000Ch
  loc_00E02576: push 00402836h ; __vbaExceptHandler
  loc_00E0257B: mov eax, fs:[00000000h]
  loc_00E02581: push eax
  loc_00E02582: mov fs:[00000000h], esp
  loc_00E02589: sub esp, 00000028h
  loc_00E0258C: push ebx
  loc_00E0258D: push esi
  loc_00E0258E: push edi
  loc_00E0258F: mov var_C, esp
  loc_00E02592: mov var_8, 00401D80h
  loc_00E02599: xor ebx, ebx
  loc_00E0259B: mov var_4, ebx
  loc_00E0259E: mov esi, Me
  loc_00E025A1: push esi
  loc_00E025A2: mov eax, [esi]
  loc_00E025A4: call [eax+00000004h]
  loc_00E025A7: mov ecx, New_Picture
  loc_00E025AA: mov edi, [004010B4h] ; __vbaObjSetAddref
  loc_00E025B0: lea edx, var_18
  loc_00E025B3: push ecx
  loc_00E025B4: push edx
  loc_00E025B5: mov var_18, ebx
  loc_00E025B8: mov var_1C, ebx
  loc_00E025BB: call edi
  loc_00E025BD: mov eax, var_18
  loc_00E025C0: lea ecx, [esi+0000014Ch]
  loc_00E025C6: push eax
  loc_00E025C7: push ecx
  loc_00E025C8: call edi
  loc_00E025CA: mov edx, var_18
  loc_00E025CD: push ebx
  loc_00E025CE: push edx
  loc_00E025CF: call [0040114Ch] ; __vbaObjIs
  loc_00E025D5: test ax, ax
  loc_00E025D8: jnz 00E0261Bh
  loc_00E025DA: mov eax, [esi]
  loc_00E025DC: push esi
  loc_00E025DD: call [eax+00000910h]
  loc_00E025E3: sub esp, 00000010h
  loc_00E025E6: mov ecx, 00000008h
  loc_00E025EB: mov edi, esp
  loc_00E025ED: mov edx, [esi]
  loc_00E025EF: mov eax, 006DEC6Ch ; "PictureNormal"
  loc_00E025F4: push esi
  loc_00E025F5: mov [edi], ecx
  loc_00E025F7: mov ecx, var_28
  loc_00E025FA: mov [edi+00000004h], ecx
  loc_00E025FD: mov [edi+00000008h], eax
  loc_00E02600: mov eax, var_20
  loc_00E02603: mov [edi+0000000Ch], eax
  loc_00E02606: call [edx+00000390h]
  loc_00E0260C: cmp eax, ebx
  loc_00E0260E: fnclex
  loc_00E02610: jge 00E0270Fh
  loc_00E02616: jmp 00E026FDh
  loc_00E0261B: mov ecx, [esi]
  loc_00E0261D: push esi
  loc_00E0261E: call [ecx+000009C4h]
  loc_00E02624: cmp eax, ebx
  loc_00E02626: jge 00E0263Ah
  loc_00E02628: push 000009C4h
  loc_00E0262D: push 006DD090h
  loc_00E02632: push esi
  loc_00E02633: push eax
  loc_00E02634: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0263A: push 006DE764h
  loc_00E0263F: push ebx
  loc_00E02640: mov ebx, [00401224h] ; __vbaCastObj
  loc_00E02646: call ebx
  loc_00E02648: lea edx, var_1C
  loc_00E0264B: push eax
  loc_00E0264C: push edx
  loc_00E0264D: call [004010ACh] ; __vbaObjSet
  loc_00E02653: push eax
  loc_00E02654: lea eax, [esi+00000150h]
  loc_00E0265A: push eax
  loc_00E0265B: call edi
  loc_00E0265D: lea ecx, var_1C
  loc_00E02660: call [00401254h] ; __vbaFreeObj
  loc_00E02666: push 006DE764h
  loc_00E0266B: push 00000000h
  loc_00E0266D: call ebx
  loc_00E0266F: lea ecx, var_1C
  loc_00E02672: push eax
  loc_00E02673: push ecx
  loc_00E02674: call [004010ACh] ; __vbaObjSet
  loc_00E0267A: lea edx, [esi+00000154h]
  loc_00E02680: push eax
  loc_00E02681: push edx
  loc_00E02682: call edi
  loc_00E02684: lea ecx, var_1C
  loc_00E02687: call [00401254h] ; __vbaFreeObj
  loc_00E0268D: mov edx, [esi]
  loc_00E0268F: mov edi, var_28
  loc_00E02692: sub esp, 00000010h
  loc_00E02695: mov var_3C, edx
  loc_00E02698: mov edx, esp
  loc_00E0269A: mov ecx, 00000008h
  loc_00E0269F: mov ebx, var_20
  loc_00E026A2: mov eax, 006DEC8Ch ; "PictureHot"
  loc_00E026A7: mov [edx], ecx
  loc_00E026A9: push esi
  loc_00E026AA: mov [edx+00000004h], edi
  loc_00E026AD: mov [edx+00000008h], eax
  loc_00E026B0: mov eax, var_3C
  loc_00E026B3: mov [edx+0000000Ch], ebx
  loc_00E026B6: call [eax+00000390h]
  loc_00E026BC: test eax, eax
  loc_00E026BE: fnclex
  loc_00E026C0: jge 00E026D4h
  loc_00E026C2: push 00000390h
  loc_00E026C7: push 006DCB00h
  loc_00E026CC: push esi
  loc_00E026CD: push eax
  loc_00E026CE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E026D4: sub esp, 00000010h
  loc_00E026D7: mov eax, 00000008h
  loc_00E026DC: mov edx, esp
  loc_00E026DE: mov ecx, [esi]
  loc_00E026E0: push esi
  loc_00E026E1: mov [edx], eax
  loc_00E026E3: mov eax, 006DECA8h ; "PictureDown"
  loc_00E026E8: mov [edx+00000004h], edi
  loc_00E026EB: mov [edx+00000008h], eax
  loc_00E026EE: mov [edx+0000000Ch], ebx
  loc_00E026F1: call [ecx+00000390h]
  loc_00E026F7: test eax, eax
  loc_00E026F9: fnclex
  loc_00E026FB: jge 00E0270Fh
  loc_00E026FD: push 00000390h
  loc_00E02702: push 006DCB00h
  loc_00E02707: push esi
  loc_00E02708: push eax
  loc_00E02709: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0270F: push 00E0272Ah
  loc_00E02714: jmp 00E02720h
  loc_00E02716: lea ecx, var_1C
  loc_00E02719: call [00401254h] ; __vbaFreeObj
  loc_00E0271F: ret
  loc_00E02720: lea ecx, var_18
  loc_00E02723: call [00401254h] ; __vbaFreeObj
  loc_00E02729: ret
  loc_00E0272A: mov eax, Me
  loc_00E0272D: push eax
  loc_00E0272E: mov ecx, [eax]
  loc_00E02730: call [ecx+00000008h]
  loc_00E02733: mov eax, var_4
  loc_00E02736: mov ecx, var_14
  loc_00E02739: pop edi
  loc_00E0273A: pop esi
  loc_00E0273B: mov fs:[00000000h], ecx
  loc_00E02742: pop ebx
  loc_00E02743: mov esp, ebp
  loc_00E02745: pop ebp
  loc_00E02746: retn 0008h
End Sub

Public Property Get PictureHot(arg_C) 'E02750
  loc_00E02750: push ebp
  loc_00E02751: mov ebp, esp
  loc_00E02753: sub esp, 0000000Ch
  loc_00E02756: push 00402836h ; __vbaExceptHandler
  loc_00E0275B: mov eax, fs:[00000000h]
  loc_00E02761: push eax
  loc_00E02762: mov fs:[00000000h], esp
  loc_00E02769: sub esp, 0000000Ch
  loc_00E0276C: push ebx
  loc_00E0276D: push esi
  loc_00E0276E: push edi
  loc_00E0276F: mov var_C, esp
  loc_00E02772: mov var_8, 00401D90h
  loc_00E02779: xor edi, edi
  loc_00E0277B: mov var_4, edi
  loc_00E0277E: mov esi, Me
  loc_00E02781: push esi
  loc_00E02782: mov eax, [esi]
  loc_00E02784: call [eax+00000004h]
  loc_00E02787: mov ecx, arg_C
  loc_00E0278A: lea eax, var_18
  loc_00E0278D: mov var_18, edi
  loc_00E02790: mov [ecx], edi
  loc_00E02792: mov edx, [esi+00000150h]
  loc_00E02798: push edx
  loc_00E02799: push eax
  loc_00E0279A: call [004010B4h] ; __vbaObjSetAddref
  loc_00E027A0: push 00E027B2h
  loc_00E027A5: jmp 00E027B1h
  loc_00E027A7: lea ecx, var_18
  loc_00E027AA: call [00401254h] ; __vbaFreeObj
  loc_00E027B0: ret
  loc_00E027B1: ret
  loc_00E027B2: mov eax, Me
  loc_00E027B5: push eax
  loc_00E027B6: mov ecx, [eax]
  loc_00E027B8: call [ecx+00000008h]
  loc_00E027BB: mov edx, arg_C
  loc_00E027BE: mov eax, var_18
  loc_00E027C1: mov [edx], eax
  loc_00E027C3: mov eax, var_4
  loc_00E027C6: mov ecx, var_14
  loc_00E027C9: pop edi
  loc_00E027CA: pop esi
  loc_00E027CB: mov fs:[00000000h], ecx
  loc_00E027D2: pop ebx
  loc_00E027D3: mov esp, ebp
  loc_00E027D5: pop ebp
  loc_00E027D6: retn 0008h
End Sub

Public Property Set PictureHot(New_Hot) 'E027E0
  loc_00E027E0: push ebp
  loc_00E027E1: mov ebp, esp
  loc_00E027E3: sub esp, 0000000Ch
  loc_00E027E6: push 00402836h ; __vbaExceptHandler
  loc_00E027EB: mov eax, fs:[00000000h]
  loc_00E027F1: push eax
  loc_00E027F2: mov fs:[00000000h], esp
  loc_00E027F9: sub esp, 00000020h
  loc_00E027FC: push ebx
  loc_00E027FD: push esi
  loc_00E027FE: push edi
  loc_00E027FF: mov var_C, esp
  loc_00E02802: mov var_8, 00401DA0h
  loc_00E02809: xor edi, edi
  loc_00E0280B: mov var_4, edi
  loc_00E0280E: mov esi, Me
  loc_00E02811: push esi
  loc_00E02812: mov eax, [esi]
  loc_00E02814: call [eax+00000004h]
  loc_00E02817: mov ecx, New_Hot
  loc_00E0281A: lea edx, var_18
  loc_00E0281D: mov var_18, edi
  loc_00E02820: mov edi, [004010B4h] ; __vbaObjSetAddref
  loc_00E02826: push ecx
  loc_00E02827: push edx
  loc_00E02828: call edi
  loc_00E0282A: mov eax, [esi+0000014Ch]
  loc_00E02830: lea ebx, [esi+0000014Ch]
  loc_00E02836: push 00000000h
  loc_00E02838: push eax
  loc_00E02839: call [0040114Ch] ; __vbaObjIs
  loc_00E0283F: mov ecx, var_18
  loc_00E02842: test ax, ax
  loc_00E02845: push ecx
  loc_00E02846: jz 00E0288Eh
  loc_00E02848: push ebx
  loc_00E02849: call edi
  loc_00E0284B: sub esp, 00000010h
  loc_00E0284E: mov ecx, 00000008h
  loc_00E02853: mov edi, esp
  loc_00E02855: mov edx, [esi]
  loc_00E02857: mov eax, 006DEC6Ch ; "PictureNormal"
  loc_00E0285C: push esi
  loc_00E0285D: mov [edi], ecx
  loc_00E0285F: mov ecx, var_24
  loc_00E02862: mov [edi+00000004h], ecx
  loc_00E02865: mov [edi+00000008h], eax
  loc_00E02868: mov eax, var_1C
  loc_00E0286B: mov [edi+0000000Ch], eax
  loc_00E0286E: call [edx+00000390h]
  loc_00E02874: test eax, eax
  loc_00E02876: fnclex
  loc_00E02878: jge 00E028E1h
  loc_00E0287A: push 00000390h
  loc_00E0287F: push 006DCB00h
  loc_00E02884: push esi
  loc_00E02885: push eax
  loc_00E02886: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0288C: jmp 00E028E1h
  loc_00E0288E: lea edx, [esi+00000150h]
  loc_00E02894: push edx
  loc_00E02895: call edi
  loc_00E02897: sub esp, 00000010h
  loc_00E0289A: mov ecx, 00000008h
  loc_00E0289F: mov edi, esp
  loc_00E028A1: mov edx, [esi]
  loc_00E028A3: mov eax, 006DEC8Ch ; "PictureHot"
  loc_00E028A8: push esi
  loc_00E028A9: mov [edi], ecx
  loc_00E028AB: mov ecx, var_24
  loc_00E028AE: mov [edi+00000004h], ecx
  loc_00E028B1: mov [edi+00000008h], eax
  loc_00E028B4: mov eax, var_1C
  loc_00E028B7: mov [edi+0000000Ch], eax
  loc_00E028BA: call [edx+00000390h]
  loc_00E028C0: test eax, eax
  loc_00E028C2: fnclex
  loc_00E028C4: jge 00E028D8h
  loc_00E028C6: push 00000390h
  loc_00E028CB: push 006DCB00h
  loc_00E028D0: push esi
  loc_00E028D1: push eax
  loc_00E028D2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E028D8: mov ecx, [esi]
  loc_00E028DA: push esi
  loc_00E028DB: call [ecx+00000910h]
  loc_00E028E1: push 00E028F0h
  loc_00E028E6: lea ecx, var_18
  loc_00E028E9: call [00401254h] ; __vbaFreeObj
  loc_00E028EF: ret
  loc_00E028F0: mov eax, Me
  loc_00E028F3: push eax
  loc_00E028F4: mov edx, [eax]
  loc_00E028F6: call [edx+00000008h]
  loc_00E028F9: mov eax, var_4
  loc_00E028FC: mov ecx, var_14
  loc_00E028FF: pop edi
  loc_00E02900: pop esi
  loc_00E02901: mov fs:[00000000h], ecx
  loc_00E02908: pop ebx
  loc_00E02909: mov esp, ebp
  loc_00E0290B: pop ebp
  loc_00E0290C: retn 0008h
End Sub

Public Property Get PictureDown(arg_C) 'E02910
  loc_00E02910: push ebp
  loc_00E02911: mov ebp, esp
  loc_00E02913: sub esp, 0000000Ch
  loc_00E02916: push 00402836h ; __vbaExceptHandler
  loc_00E0291B: mov eax, fs:[00000000h]
  loc_00E02921: push eax
  loc_00E02922: mov fs:[00000000h], esp
  loc_00E02929: sub esp, 0000000Ch
  loc_00E0292C: push ebx
  loc_00E0292D: push esi
  loc_00E0292E: push edi
  loc_00E0292F: mov var_C, esp
  loc_00E02932: mov var_8, 00401DB0h
  loc_00E02939: xor edi, edi
  loc_00E0293B: mov var_4, edi
  loc_00E0293E: mov esi, Me
  loc_00E02941: push esi
  loc_00E02942: mov eax, [esi]
  loc_00E02944: call [eax+00000004h]
  loc_00E02947: mov ecx, arg_C
  loc_00E0294A: lea eax, var_18
  loc_00E0294D: mov var_18, edi
  loc_00E02950: mov [ecx], edi
  loc_00E02952: mov edx, [esi+00000154h]
  loc_00E02958: push edx
  loc_00E02959: push eax
  loc_00E0295A: call [004010B4h] ; __vbaObjSetAddref
  loc_00E02960: push 00E02972h
  loc_00E02965: jmp 00E02971h
  loc_00E02967: lea ecx, var_18
  loc_00E0296A: call [00401254h] ; __vbaFreeObj
  loc_00E02970: ret
  loc_00E02971: ret
  loc_00E02972: mov eax, Me
  loc_00E02975: push eax
  loc_00E02976: mov ecx, [eax]
  loc_00E02978: call [ecx+00000008h]
  loc_00E0297B: mov edx, arg_C
  loc_00E0297E: mov eax, var_18
  loc_00E02981: mov [edx], eax
  loc_00E02983: mov eax, var_4
  loc_00E02986: mov ecx, var_14
  loc_00E02989: pop edi
  loc_00E0298A: pop esi
  loc_00E0298B: mov fs:[00000000h], ecx
  loc_00E02992: pop ebx
  loc_00E02993: mov esp, ebp
  loc_00E02995: pop ebp
  loc_00E02996: retn 0008h
End Sub

Public Property Set PictureDown(New_Down) 'E029A0
  loc_00E029A0: push ebp
  loc_00E029A1: mov ebp, esp
  loc_00E029A3: sub esp, 0000000Ch
  loc_00E029A6: push 00402836h ; __vbaExceptHandler
  loc_00E029AB: mov eax, fs:[00000000h]
  loc_00E029B1: push eax
  loc_00E029B2: mov fs:[00000000h], esp
  loc_00E029B9: sub esp, 00000020h
  loc_00E029BC: push ebx
  loc_00E029BD: push esi
  loc_00E029BE: push edi
  loc_00E029BF: mov var_C, esp
  loc_00E029C2: mov var_8, 00401DC0h
  loc_00E029C9: xor edi, edi
  loc_00E029CB: mov var_4, edi
  loc_00E029CE: mov esi, Me
  loc_00E029D1: push esi
  loc_00E029D2: mov eax, [esi]
  loc_00E029D4: call [eax+00000004h]
  loc_00E029D7: mov ecx, New_Down
  loc_00E029DA: lea edx, var_18
  loc_00E029DD: mov var_18, edi
  loc_00E029E0: mov edi, [004010B4h] ; __vbaObjSetAddref
  loc_00E029E6: push ecx
  loc_00E029E7: push edx
  loc_00E029E8: call edi
  loc_00E029EA: mov eax, [esi+0000014Ch]
  loc_00E029F0: lea ebx, [esi+0000014Ch]
  loc_00E029F6: push 00000000h
  loc_00E029F8: push eax
  loc_00E029F9: call [0040114Ch] ; __vbaObjIs
  loc_00E029FF: mov ecx, var_18
  loc_00E02A02: test ax, ax
  loc_00E02A05: push ecx
  loc_00E02A06: jz 00E02A4Eh
  loc_00E02A08: push ebx
  loc_00E02A09: call edi
  loc_00E02A0B: sub esp, 00000010h
  loc_00E02A0E: mov ecx, 00000008h
  loc_00E02A13: mov edi, esp
  loc_00E02A15: mov edx, [esi]
  loc_00E02A17: mov eax, 006DEC6Ch ; "PictureNormal"
  loc_00E02A1C: push esi
  loc_00E02A1D: mov [edi], ecx
  loc_00E02A1F: mov ecx, var_24
  loc_00E02A22: mov [edi+00000004h], ecx
  loc_00E02A25: mov [edi+00000008h], eax
  loc_00E02A28: mov eax, var_1C
  loc_00E02A2B: mov [edi+0000000Ch], eax
  loc_00E02A2E: call [edx+00000390h]
  loc_00E02A34: test eax, eax
  loc_00E02A36: fnclex
  loc_00E02A38: jge 00E02AA1h
  loc_00E02A3A: push 00000390h
  loc_00E02A3F: push 006DCB00h
  loc_00E02A44: push esi
  loc_00E02A45: push eax
  loc_00E02A46: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E02A4C: jmp 00E02AA1h
  loc_00E02A4E: lea edx, [esi+00000154h]
  loc_00E02A54: push edx
  loc_00E02A55: call edi
  loc_00E02A57: sub esp, 00000010h
  loc_00E02A5A: mov ecx, 00000008h
  loc_00E02A5F: mov edi, esp
  loc_00E02A61: mov edx, [esi]
  loc_00E02A63: mov eax, 006DECA8h ; "PictureDown"
  loc_00E02A68: push esi
  loc_00E02A69: mov [edi], ecx
  loc_00E02A6B: mov ecx, var_24
  loc_00E02A6E: mov [edi+00000004h], ecx
  loc_00E02A71: mov [edi+00000008h], eax
  loc_00E02A74: mov eax, var_1C
  loc_00E02A77: mov [edi+0000000Ch], eax
  loc_00E02A7A: call [edx+00000390h]
  loc_00E02A80: test eax, eax
  loc_00E02A82: fnclex
  loc_00E02A84: jge 00E02A98h
  loc_00E02A86: push 00000390h
  loc_00E02A8B: push 006DCB00h
  loc_00E02A90: push esi
  loc_00E02A91: push eax
  loc_00E02A92: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E02A98: mov ecx, [esi]
  loc_00E02A9A: push esi
  loc_00E02A9B: call [ecx+00000910h]
  loc_00E02AA1: push 00E02AB0h
  loc_00E02AA6: lea ecx, var_18
  loc_00E02AA9: call [00401254h] ; __vbaFreeObj
  loc_00E02AAF: ret
  loc_00E02AB0: mov eax, Me
  loc_00E02AB3: push eax
  loc_00E02AB4: mov edx, [eax]
  loc_00E02AB6: call [edx+00000008h]
  loc_00E02AB9: mov eax, var_4
  loc_00E02ABC: mov ecx, var_14
  loc_00E02ABF: pop edi
  loc_00E02AC0: pop esi
  loc_00E02AC1: mov fs:[00000000h], ecx
  loc_00E02AC8: pop ebx
  loc_00E02AC9: mov esp, ebp
  loc_00E02ACB: pop ebp
  loc_00E02ACC: retn 0008h
End Sub

Public Property Get PictureAlign(arg_C) 'E02AD0
  loc_00E02AD0: push ebp
  loc_00E02AD1: mov ebp, esp
  loc_00E02AD3: sub esp, 0000000Ch
  loc_00E02AD6: push 00402836h ; __vbaExceptHandler
  loc_00E02ADB: mov eax, fs:[00000000h]
  loc_00E02AE1: push eax
  loc_00E02AE2: mov fs:[00000000h], esp
  loc_00E02AE9: sub esp, 0000000Ch
  loc_00E02AEC: push ebx
  loc_00E02AED: push esi
  loc_00E02AEE: push edi
  loc_00E02AEF: mov var_C, esp
  loc_00E02AF2: mov var_8, 00401DD0h
  loc_00E02AF9: xor edi, edi
  loc_00E02AFB: mov var_4, edi
  loc_00E02AFE: mov esi, Me
  loc_00E02B01: push esi
  loc_00E02B02: mov eax, [esi]
  loc_00E02B04: call [eax+00000004h]
  loc_00E02B07: mov ecx, [esi+00000164h]
  loc_00E02B0D: mov var_18, edi
  loc_00E02B10: mov var_18, ecx
  loc_00E02B13: mov eax, Me
  loc_00E02B16: push eax
  loc_00E02B17: mov edx, [eax]
  loc_00E02B19: call [edx+00000008h]
  loc_00E02B1C: mov eax, arg_C
  loc_00E02B1F: mov ecx, var_18
  loc_00E02B22: mov [eax], ecx
  loc_00E02B24: mov eax, var_4
  loc_00E02B27: mov ecx, var_14
  loc_00E02B2A: pop edi
  loc_00E02B2B: pop esi
  loc_00E02B2C: mov fs:[00000000h], ecx
  loc_00E02B33: pop ebx
  loc_00E02B34: mov esp, ebp
  loc_00E02B36: pop ebp
  loc_00E02B37: retn 0008h
End Sub

Public Property Let PictureAlign(New_PictureAlign) 'E02B40
  loc_00E02B40: push ebp
  loc_00E02B41: mov ebp, esp
  loc_00E02B43: sub esp, 0000000Ch
  loc_00E02B46: push 00402836h ; __vbaExceptHandler
  loc_00E02B4B: mov eax, fs:[00000000h]
  loc_00E02B51: push eax
  loc_00E02B52: mov fs:[00000000h], esp
  loc_00E02B59: sub esp, 0000001Ch
  loc_00E02B5C: push ebx
  loc_00E02B5D: push esi
  loc_00E02B5E: push edi
  loc_00E02B5F: mov var_C, esp
  loc_00E02B62: mov var_8, 00401DD8h
  loc_00E02B69: mov var_4, 00000000h
  loc_00E02B70: mov esi, Me
  loc_00E02B73: push esi
  loc_00E02B74: mov eax, [esi]
  loc_00E02B76: call [eax+00000004h]
  loc_00E02B79: mov edx, [esi+0000014Ch]
  loc_00E02B7F: mov ecx, New_PictureAlign
  loc_00E02B82: push 00000000h
  loc_00E02B84: push edx
  loc_00E02B85: mov [esi+00000164h], ecx
  loc_00E02B8B: call [0040114Ch] ; __vbaObjIs
  loc_00E02B91: test ax, ax
  loc_00E02B94: jnz 00E02B9Fh
  loc_00E02B96: mov eax, [esi]
  loc_00E02B98: push esi
  loc_00E02B99: call [eax+00000910h]
  loc_00E02B9F: sub esp, 00000010h
  loc_00E02BA2: mov ecx, 00000008h
  loc_00E02BA7: mov edi, esp
  loc_00E02BA9: mov edx, [esi]
  loc_00E02BAB: mov eax, 006DEE54h ; "PictureAlign"
  loc_00E02BB0: push esi
  loc_00E02BB1: mov [edi], ecx
  loc_00E02BB3: mov ecx, var_20
  loc_00E02BB6: mov [edi+00000004h], ecx
  loc_00E02BB9: mov [edi+00000008h], eax
  loc_00E02BBC: mov eax, var_18
  loc_00E02BBF: mov [edi+0000000Ch], eax
  loc_00E02BC2: call [edx+00000390h]
  loc_00E02BC8: test eax, eax
  loc_00E02BCA: fnclex
  loc_00E02BCC: jge 00E02BE0h
  loc_00E02BCE: push 00000390h
  loc_00E02BD3: push 006DCB00h
  loc_00E02BD8: push esi
  loc_00E02BD9: push eax
  loc_00E02BDA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E02BE0: mov eax, Me
  loc_00E02BE3: push eax
  loc_00E02BE4: mov ecx, [eax]
  loc_00E02BE6: call [ecx+00000008h]
  loc_00E02BE9: mov eax, var_4
  loc_00E02BEC: mov ecx, var_14
  loc_00E02BEF: pop edi
  loc_00E02BF0: pop esi
  loc_00E02BF1: mov fs:[00000000h], ecx
  loc_00E02BF8: pop ebx
  loc_00E02BF9: mov esp, ebp
  loc_00E02BFB: pop ebp
  loc_00E02BFC: retn 0008h
End Sub

Public Property Get PictureShadow(arg_C) 'E02C00
  loc_00E02C00: push ebp
  loc_00E02C01: mov ebp, esp
  loc_00E02C03: sub esp, 0000000Ch
  loc_00E02C06: push 00402836h ; __vbaExceptHandler
  loc_00E02C0B: mov eax, fs:[00000000h]
  loc_00E02C11: push eax
  loc_00E02C12: mov fs:[00000000h], esp
  loc_00E02C19: sub esp, 0000000Ch
  loc_00E02C1C: push ebx
  loc_00E02C1D: push esi
  loc_00E02C1E: push edi
  loc_00E02C1F: mov var_C, esp
  loc_00E02C22: mov var_8, 00401DE0h
  loc_00E02C29: xor edi, edi
  loc_00E02C2B: mov var_4, edi
  loc_00E02C2E: mov esi, Me
  loc_00E02C31: push esi
  loc_00E02C32: mov eax, [esi]
  loc_00E02C34: call [eax+00000004h]
  loc_00E02C37: mov cx, [esi+00000160h]
  loc_00E02C3E: mov var_18, edi
  loc_00E02C41: mov var_18, ecx
  loc_00E02C44: mov eax, Me
  loc_00E02C47: push eax
  loc_00E02C48: mov edx, [eax]
  loc_00E02C4A: call [edx+00000008h]
  loc_00E02C4D: mov eax, arg_C
  loc_00E02C50: mov cx, var_18
  loc_00E02C54: mov [eax], cx
  loc_00E02C57: mov eax, var_4
  loc_00E02C5A: mov ecx, var_14
  loc_00E02C5D: pop edi
  loc_00E02C5E: pop esi
  loc_00E02C5F: mov fs:[00000000h], ecx
  loc_00E02C66: pop ebx
  loc_00E02C67: mov esp, ebp
  loc_00E02C69: pop ebp
  loc_00E02C6A: retn 0008h
End Sub

Public Property Let PictureShadow(New_Shadow) 'E02C70
  loc_00E02C70: push ebp
  loc_00E02C71: mov ebp, esp
  loc_00E02C73: sub esp, 0000000Ch
  loc_00E02C76: push 00402836h ; __vbaExceptHandler
  loc_00E02C7B: mov eax, fs:[00000000h]
  loc_00E02C81: push eax
  loc_00E02C82: mov fs:[00000000h], esp
  loc_00E02C89: sub esp, 0000001Ch
  loc_00E02C8C: push ebx
  loc_00E02C8D: push esi
  loc_00E02C8E: push edi
  loc_00E02C8F: mov var_C, esp
  loc_00E02C92: mov var_8, 00401DE8h
  loc_00E02C99: mov var_4, 00000000h
  loc_00E02CA0: mov esi, Me
  loc_00E02CA3: push esi
  loc_00E02CA4: mov eax, [esi]
  loc_00E02CA6: call [eax+00000004h]
  loc_00E02CA9: mov cx, New_Shadow
  loc_00E02CAD: mov edx, [esi]
  loc_00E02CAF: push esi
  loc_00E02CB0: mov [esi+00000160h], cx
  loc_00E02CB7: call [edx+00000910h]
  loc_00E02CBD: sub esp, 00000010h
  loc_00E02CC0: mov ecx, 00000008h
  loc_00E02CC5: mov edi, esp
  loc_00E02CC7: mov edx, [esi]
  loc_00E02CC9: mov eax, 006DECC4h ; "PictureShadow"
  loc_00E02CCE: push esi
  loc_00E02CCF: mov [edi], ecx
  loc_00E02CD1: mov ecx, var_20
  loc_00E02CD4: mov [edi+00000004h], ecx
  loc_00E02CD7: mov [edi+00000008h], eax
  loc_00E02CDA: mov eax, var_18
  loc_00E02CDD: mov [edi+0000000Ch], eax
  loc_00E02CE0: call [edx+00000390h]
  loc_00E02CE6: test eax, eax
  loc_00E02CE8: fnclex
  loc_00E02CEA: jge 00E02CFEh
  loc_00E02CEC: push 00000390h
  loc_00E02CF1: push 006DCB00h
  loc_00E02CF6: push esi
  loc_00E02CF7: push eax
  loc_00E02CF8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E02CFE: mov eax, Me
  loc_00E02D01: push eax
  loc_00E02D02: mov ecx, [eax]
  loc_00E02D04: call [ecx+00000008h]
  loc_00E02D07: mov eax, var_4
  loc_00E02D0A: mov ecx, var_14
  loc_00E02D0D: pop edi
  loc_00E02D0E: pop esi
  loc_00E02D0F: mov fs:[00000000h], ecx
  loc_00E02D16: pop ebx
  loc_00E02D17: mov esp, ebp
  loc_00E02D19: pop ebp
  loc_00E02D1A: retn 0008h
End Sub

Public Property Get PictureEffectOnOver(arg_C) 'E02D20
  loc_00E02D20: push ebp
  loc_00E02D21: mov ebp, esp
  loc_00E02D23: sub esp, 0000000Ch
  loc_00E02D26: push 00402836h ; __vbaExceptHandler
  loc_00E02D2B: mov eax, fs:[00000000h]
  loc_00E02D31: push eax
  loc_00E02D32: mov fs:[00000000h], esp
  loc_00E02D39: sub esp, 0000000Ch
  loc_00E02D3C: push ebx
  loc_00E02D3D: push esi
  loc_00E02D3E: push edi
  loc_00E02D3F: mov var_C, esp
  loc_00E02D42: mov var_8, 00401DF0h
  loc_00E02D49: xor edi, edi
  loc_00E02D4B: mov var_4, edi
  loc_00E02D4E: mov esi, Me
  loc_00E02D51: push esi
  loc_00E02D52: mov eax, [esi]
  loc_00E02D54: call [eax+00000004h]
  loc_00E02D57: mov ecx, [esi+00000168h]
  loc_00E02D5D: mov var_18, edi
  loc_00E02D60: mov var_18, ecx
  loc_00E02D63: mov eax, Me
  loc_00E02D66: push eax
  loc_00E02D67: mov edx, [eax]
  loc_00E02D69: call [edx+00000008h]
  loc_00E02D6C: mov eax, arg_C
  loc_00E02D6F: mov ecx, var_18
  loc_00E02D72: mov [eax], ecx
  loc_00E02D74: mov eax, var_4
  loc_00E02D77: mov ecx, var_14
  loc_00E02D7A: pop edi
  loc_00E02D7B: pop esi
  loc_00E02D7C: mov fs:[00000000h], ecx
  loc_00E02D83: pop ebx
  loc_00E02D84: mov esp, ebp
  loc_00E02D86: pop ebp
  loc_00E02D87: retn 0008h
End Sub

Public Property Let PictureEffectOnOver(New_Effect) 'E02D90
  loc_00E02D90: push ebp
  loc_00E02D91: mov ebp, esp
  loc_00E02D93: sub esp, 0000000Ch
  loc_00E02D96: push 00402836h ; __vbaExceptHandler
  loc_00E02D9B: mov eax, fs:[00000000h]
  loc_00E02DA1: push eax
  loc_00E02DA2: mov fs:[00000000h], esp
  loc_00E02DA9: sub esp, 0000001Ch
  loc_00E02DAC: push ebx
  loc_00E02DAD: push esi
  loc_00E02DAE: push edi
  loc_00E02DAF: mov var_C, esp
  loc_00E02DB2: mov var_8, 00401DF8h
  loc_00E02DB9: mov var_4, 00000000h
  loc_00E02DC0: mov esi, Me
  loc_00E02DC3: push esi
  loc_00E02DC4: mov eax, [esi]
  loc_00E02DC6: call [eax+00000004h]
  loc_00E02DC9: mov ecx, New_Effect
  loc_00E02DCC: mov edx, [esi]
  loc_00E02DCE: push esi
  loc_00E02DCF: mov [esi+00000168h], ecx
  loc_00E02DD5: call [edx+00000910h]
  loc_00E02DDB: sub esp, 00000010h
  loc_00E02DDE: mov ecx, 00000008h
  loc_00E02DE3: mov edi, esp
  loc_00E02DE5: mov edx, [esi]
  loc_00E02DE7: mov eax, 006DED90h ; "PictureEffectOnOver"
  loc_00E02DEC: push esi
  loc_00E02DED: mov [edi], ecx
  loc_00E02DEF: mov ecx, var_20
  loc_00E02DF2: mov [edi+00000004h], ecx
  loc_00E02DF5: mov [edi+00000008h], eax
  loc_00E02DF8: mov eax, var_18
  loc_00E02DFB: mov [edi+0000000Ch], eax
  loc_00E02DFE: call [edx+00000390h]
  loc_00E02E04: test eax, eax
  loc_00E02E06: fnclex
  loc_00E02E08: jge 00E02E1Ch
  loc_00E02E0A: push 00000390h
  loc_00E02E0F: push 006DCB00h
  loc_00E02E14: push esi
  loc_00E02E15: push eax
  loc_00E02E16: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E02E1C: mov eax, Me
  loc_00E02E1F: push eax
  loc_00E02E20: mov ecx, [eax]
  loc_00E02E22: call [ecx+00000008h]
  loc_00E02E25: mov eax, var_4
  loc_00E02E28: mov ecx, var_14
  loc_00E02E2B: pop edi
  loc_00E02E2C: pop esi
  loc_00E02E2D: mov fs:[00000000h], ecx
  loc_00E02E34: pop ebx
  loc_00E02E35: mov esp, ebp
  loc_00E02E37: pop ebp
  loc_00E02E38: retn 0008h
End Sub

Public Property Get PictureEffectOnDown(arg_C) 'E02E40
  loc_00E02E40: push ebp
  loc_00E02E41: mov ebp, esp
  loc_00E02E43: sub esp, 0000000Ch
  loc_00E02E46: push 00402836h ; __vbaExceptHandler
  loc_00E02E4B: mov eax, fs:[00000000h]
  loc_00E02E51: push eax
  loc_00E02E52: mov fs:[00000000h], esp
  loc_00E02E59: sub esp, 0000000Ch
  loc_00E02E5C: push ebx
  loc_00E02E5D: push esi
  loc_00E02E5E: push edi
  loc_00E02E5F: mov var_C, esp
  loc_00E02E62: mov var_8, 00401E00h
  loc_00E02E69: xor edi, edi
  loc_00E02E6B: mov var_4, edi
  loc_00E02E6E: mov esi, Me
  loc_00E02E71: push esi
  loc_00E02E72: mov eax, [esi]
  loc_00E02E74: call [eax+00000004h]
  loc_00E02E77: mov ecx, [esi+0000016Ch]
  loc_00E02E7D: mov var_18, edi
  loc_00E02E80: mov var_18, ecx
  loc_00E02E83: mov eax, Me
  loc_00E02E86: push eax
  loc_00E02E87: mov edx, [eax]
  loc_00E02E89: call [edx+00000008h]
  loc_00E02E8C: mov eax, arg_C
  loc_00E02E8F: mov ecx, var_18
  loc_00E02E92: mov [eax], ecx
  loc_00E02E94: mov eax, var_4
  loc_00E02E97: mov ecx, var_14
  loc_00E02E9A: pop edi
  loc_00E02E9B: pop esi
  loc_00E02E9C: mov fs:[00000000h], ecx
  loc_00E02EA3: pop ebx
  loc_00E02EA4: mov esp, ebp
  loc_00E02EA6: pop ebp
  loc_00E02EA7: retn 0008h
End Sub

Public Property Let PictureEffectOnDown(New_Effect) 'E02EB0
  loc_00E02EB0: push ebp
  loc_00E02EB1: mov ebp, esp
  loc_00E02EB3: sub esp, 0000000Ch
  loc_00E02EB6: push 00402836h ; __vbaExceptHandler
  loc_00E02EBB: mov eax, fs:[00000000h]
  loc_00E02EC1: push eax
  loc_00E02EC2: mov fs:[00000000h], esp
  loc_00E02EC9: sub esp, 0000001Ch
  loc_00E02ECC: push ebx
  loc_00E02ECD: push esi
  loc_00E02ECE: push edi
  loc_00E02ECF: mov var_C, esp
  loc_00E02ED2: mov var_8, 00401E08h
  loc_00E02ED9: mov var_4, 00000000h
  loc_00E02EE0: mov esi, Me
  loc_00E02EE3: push esi
  loc_00E02EE4: mov eax, [esi]
  loc_00E02EE6: call [eax+00000004h]
  loc_00E02EE9: mov ecx, New_Effect
  loc_00E02EEC: mov edx, [esi]
  loc_00E02EEE: push esi
  loc_00E02EEF: mov [esi+0000016Ch], ecx
  loc_00E02EF5: call [edx+00000910h]
  loc_00E02EFB: sub esp, 00000010h
  loc_00E02EFE: mov ecx, 00000008h
  loc_00E02F03: mov edi, esp
  loc_00E02F05: mov edx, [esi]
  loc_00E02F07: mov eax, 006DEDBCh ; "PictureEffectOnDown"
  loc_00E02F0C: push esi
  loc_00E02F0D: mov [edi], ecx
  loc_00E02F0F: mov ecx, var_20
  loc_00E02F12: mov [edi+00000004h], ecx
  loc_00E02F15: mov [edi+00000008h], eax
  loc_00E02F18: mov eax, var_18
  loc_00E02F1B: mov [edi+0000000Ch], eax
  loc_00E02F1E: call [edx+00000390h]
  loc_00E02F24: test eax, eax
  loc_00E02F26: fnclex
  loc_00E02F28: jge 00E02F3Ch
  loc_00E02F2A: push 00000390h
  loc_00E02F2F: push 006DCB00h
  loc_00E02F34: push esi
  loc_00E02F35: push eax
  loc_00E02F36: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E02F3C: mov eax, Me
  loc_00E02F3F: push eax
  loc_00E02F40: mov ecx, [eax]
  loc_00E02F42: call [ecx+00000008h]
  loc_00E02F45: mov eax, var_4
  loc_00E02F48: mov ecx, var_14
  loc_00E02F4B: pop edi
  loc_00E02F4C: pop esi
  loc_00E02F4D: mov fs:[00000000h], ecx
  loc_00E02F54: pop ebx
  loc_00E02F55: mov esp, ebp
  loc_00E02F57: pop ebp
  loc_00E02F58: retn 0008h
End Sub

Public Property Get PicturePushOnHover(arg_C) 'E02F60
  loc_00E02F60: push ebp
  loc_00E02F61: mov ebp, esp
  loc_00E02F63: sub esp, 0000000Ch
  loc_00E02F66: push 00402836h ; __vbaExceptHandler
  loc_00E02F6B: mov eax, fs:[00000000h]
  loc_00E02F71: push eax
  loc_00E02F72: mov fs:[00000000h], esp
  loc_00E02F79: sub esp, 0000000Ch
  loc_00E02F7C: push ebx
  loc_00E02F7D: push esi
  loc_00E02F7E: push edi
  loc_00E02F7F: mov var_C, esp
  loc_00E02F82: mov var_8, 00401E10h
  loc_00E02F89: xor edi, edi
  loc_00E02F8B: mov var_4, edi
  loc_00E02F8E: mov esi, Me
  loc_00E02F91: push esi
  loc_00E02F92: mov eax, [esi]
  loc_00E02F94: call [eax+00000004h]
  loc_00E02F97: mov cx, [esi+00000170h]
  loc_00E02F9E: mov var_18, edi
  loc_00E02FA1: mov var_18, ecx
  loc_00E02FA4: mov eax, Me
  loc_00E02FA7: push eax
  loc_00E02FA8: mov edx, [eax]
  loc_00E02FAA: call [edx+00000008h]
  loc_00E02FAD: mov eax, arg_C
  loc_00E02FB0: mov cx, var_18
  loc_00E02FB4: mov [eax], cx
  loc_00E02FB7: mov eax, var_4
  loc_00E02FBA: mov ecx, var_14
  loc_00E02FBD: pop edi
  loc_00E02FBE: pop esi
  loc_00E02FBF: mov fs:[00000000h], ecx
  loc_00E02FC6: pop ebx
  loc_00E02FC7: mov esp, ebp
  loc_00E02FC9: pop ebp
  loc_00E02FCA: retn 0008h
End Sub

Public Property Let PicturePushOnHover(Value) 'E02FD0
  loc_00E02FD0: push ebp
  loc_00E02FD1: mov ebp, esp
  loc_00E02FD3: sub esp, 0000000Ch
  loc_00E02FD6: push 00402836h ; __vbaExceptHandler
  loc_00E02FDB: mov eax, fs:[00000000h]
  loc_00E02FE1: push eax
  loc_00E02FE2: mov fs:[00000000h], esp
  loc_00E02FE9: sub esp, 0000001Ch
  loc_00E02FEC: push ebx
  loc_00E02FED: push esi
  loc_00E02FEE: push edi
  loc_00E02FEF: mov var_C, esp
  loc_00E02FF2: mov var_8, 00401E18h
  loc_00E02FF9: mov var_4, 00000000h
  loc_00E03000: mov esi, Me
  loc_00E03003: push esi
  loc_00E03004: mov eax, [esi]
  loc_00E03006: call [eax+00000004h]
  loc_00E03009: mov cx, Value
  loc_00E0300D: mov edx, [esi]
  loc_00E0300F: push esi
  loc_00E03010: mov [esi+00000170h], cx
  loc_00E03017: call [edx+00000910h]
  loc_00E0301D: sub esp, 00000010h
  loc_00E03020: mov ecx, 00000008h
  loc_00E03025: mov edi, esp
  loc_00E03027: mov edx, [esi]
  loc_00E03029: mov eax, 006DED64h ; "PicturePushOnHover"
  loc_00E0302E: push esi
  loc_00E0302F: mov [edi], ecx
  loc_00E03031: mov ecx, var_20
  loc_00E03034: mov [edi+00000004h], ecx
  loc_00E03037: mov [edi+00000008h], eax
  loc_00E0303A: mov eax, var_18
  loc_00E0303D: mov [edi+0000000Ch], eax
  loc_00E03040: call [edx+00000390h]
  loc_00E03046: test eax, eax
  loc_00E03048: fnclex
  loc_00E0304A: jge 00E0305Eh
  loc_00E0304C: push 00000390h
  loc_00E03051: push 006DCB00h
  loc_00E03056: push esi
  loc_00E03057: push eax
  loc_00E03058: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0305E: mov eax, Me
  loc_00E03061: push eax
  loc_00E03062: mov ecx, [eax]
  loc_00E03064: call [ecx+00000008h]
  loc_00E03067: mov eax, var_4
  loc_00E0306A: mov ecx, var_14
  loc_00E0306D: pop edi
  loc_00E0306E: pop esi
  loc_00E0306F: mov fs:[00000000h], ecx
  loc_00E03076: pop ebx
  loc_00E03077: mov esp, ebp
  loc_00E03079: pop ebp
  loc_00E0307A: retn 0008h
End Sub

Public Property Get PictureOpacity(arg_C) 'E03080
  loc_00E03080: push ebp
  loc_00E03081: mov ebp, esp
  loc_00E03083: sub esp, 0000000Ch
  loc_00E03086: push 00402836h ; __vbaExceptHandler
  loc_00E0308B: mov eax, fs:[00000000h]
  loc_00E03091: push eax
  loc_00E03092: mov fs:[00000000h], esp
  loc_00E03099: sub esp, 0000000Ch
  loc_00E0309C: push ebx
  loc_00E0309D: push esi
  loc_00E0309E: push edi
  loc_00E0309F: mov var_C, esp
  loc_00E030A2: mov var_8, 00401E20h
  loc_00E030A9: xor ebx, ebx
  loc_00E030AB: mov var_4, ebx
  loc_00E030AE: mov esi, Me
  loc_00E030B1: push esi
  loc_00E030B2: mov eax, [esi]
  loc_00E030B4: call [eax+00000004h]
  loc_00E030B7: mov cl, [esi+00000158h]
  loc_00E030BD: mov var_18, bl
  loc_00E030C0: mov var_18, cl
  loc_00E030C3: mov eax, Me
  loc_00E030C6: push eax
  loc_00E030C7: mov edx, [eax]
  loc_00E030C9: call [edx+00000008h]
  loc_00E030CC: mov eax, arg_C
  loc_00E030CF: mov cl, var_18
  loc_00E030D2: mov [eax], cl
  loc_00E030D4: mov eax, var_4
  loc_00E030D7: mov ecx, var_14
  loc_00E030DA: pop edi
  loc_00E030DB: pop esi
  loc_00E030DC: mov fs:[00000000h], ecx
  loc_00E030E3: pop ebx
  loc_00E030E4: mov esp, ebp
  loc_00E030E6: pop ebp
  loc_00E030E7: retn 0008h
End Sub

Public Property Let PictureOpacity(New_Opacity) 'E030F0
  loc_00E030F0: push ebp
  loc_00E030F1: mov ebp, esp
  loc_00E030F3: sub esp, 0000000Ch
  loc_00E030F6: push 00402836h ; __vbaExceptHandler
  loc_00E030FB: mov eax, fs:[00000000h]
  loc_00E03101: push eax
  loc_00E03102: mov fs:[00000000h], esp
  loc_00E03109: sub esp, 0000001Ch
  loc_00E0310C: push ebx
  loc_00E0310D: push esi
  loc_00E0310E: push edi
  loc_00E0310F: mov var_C, esp
  loc_00E03112: mov var_8, 00401E28h
  loc_00E03119: mov var_4, 00000000h
  loc_00E03120: mov esi, Me
  loc_00E03123: push esi
  loc_00E03124: mov eax, [esi]
  loc_00E03126: call [eax+00000004h]
  loc_00E03129: mov cl, New_Opacity
  loc_00E0312C: mov edx, [esi]
  loc_00E0312E: push esi
  loc_00E0312F: mov [esi+00000158h], cl
  loc_00E03135: call [edx+00000910h]
  loc_00E0313B: sub esp, 00000010h
  loc_00E0313E: mov ecx, 00000008h
  loc_00E03143: mov edi, esp
  loc_00E03145: mov edx, [esi]
  loc_00E03147: mov eax, 006DECE4h ; "PictureOpacity"
  loc_00E0314C: push esi
  loc_00E0314D: mov [edi], ecx
  loc_00E0314F: mov ecx, var_20
  loc_00E03152: mov [edi+00000004h], ecx
  loc_00E03155: mov [edi+00000008h], eax
  loc_00E03158: mov eax, var_18
  loc_00E0315B: mov [edi+0000000Ch], eax
  loc_00E0315E: call [edx+00000390h]
  loc_00E03164: test eax, eax
  loc_00E03166: fnclex
  loc_00E03168: jge 00E0317Ch
  loc_00E0316A: push 00000390h
  loc_00E0316F: push 006DCB00h
  loc_00E03174: push esi
  loc_00E03175: push eax
  loc_00E03176: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0317C: mov eax, Me
  loc_00E0317F: push eax
  loc_00E03180: mov ecx, [eax]
  loc_00E03182: call [ecx+00000008h]
  loc_00E03185: mov eax, var_4
  loc_00E03188: mov ecx, var_14
  loc_00E0318B: pop edi
  loc_00E0318C: pop esi
  loc_00E0318D: mov fs:[00000000h], ecx
  loc_00E03194: pop ebx
  loc_00E03195: mov esp, ebp
  loc_00E03197: pop ebp
  loc_00E03198: retn 0008h
End Sub

Public Property Get PictureOpacityOnOver(arg_C) 'E031A0
  loc_00E031A0: push ebp
  loc_00E031A1: mov ebp, esp
  loc_00E031A3: sub esp, 0000000Ch
  loc_00E031A6: push 00402836h ; __vbaExceptHandler
  loc_00E031AB: mov eax, fs:[00000000h]
  loc_00E031B1: push eax
  loc_00E031B2: mov fs:[00000000h], esp
  loc_00E031B9: sub esp, 0000000Ch
  loc_00E031BC: push ebx
  loc_00E031BD: push esi
  loc_00E031BE: push edi
  loc_00E031BF: mov var_C, esp
  loc_00E031C2: mov var_8, 00401E30h
  loc_00E031C9: xor ebx, ebx
  loc_00E031CB: mov var_4, ebx
  loc_00E031CE: mov esi, Me
  loc_00E031D1: push esi
  loc_00E031D2: mov eax, [esi]
  loc_00E031D4: call [eax+00000004h]
  loc_00E031D7: mov cl, [esi+00000159h]
  loc_00E031DD: mov var_18, bl
  loc_00E031E0: mov var_18, cl
  loc_00E031E3: mov eax, Me
  loc_00E031E6: push eax
  loc_00E031E7: mov edx, [eax]
  loc_00E031E9: call [edx+00000008h]
  loc_00E031EC: mov eax, arg_C
  loc_00E031EF: mov cl, var_18
  loc_00E031F2: mov [eax], cl
  loc_00E031F4: mov eax, var_4
  loc_00E031F7: mov ecx, var_14
  loc_00E031FA: pop edi
  loc_00E031FB: pop esi
  loc_00E031FC: mov fs:[00000000h], ecx
  loc_00E03203: pop ebx
  loc_00E03204: mov esp, ebp
  loc_00E03206: pop ebp
  loc_00E03207: retn 0008h
End Sub

Public Property Let PictureOpacityOnOver(New_Opacity) 'E03210
  loc_00E03210: push ebp
  loc_00E03211: mov ebp, esp
  loc_00E03213: sub esp, 0000000Ch
  loc_00E03216: push 00402836h ; __vbaExceptHandler
  loc_00E0321B: mov eax, fs:[00000000h]
  loc_00E03221: push eax
  loc_00E03222: mov fs:[00000000h], esp
  loc_00E03229: sub esp, 0000001Ch
  loc_00E0322C: push ebx
  loc_00E0322D: push esi
  loc_00E0322E: push edi
  loc_00E0322F: mov var_C, esp
  loc_00E03232: mov var_8, 00401E38h
  loc_00E03239: mov var_4, 00000000h
  loc_00E03240: mov esi, Me
  loc_00E03243: push esi
  loc_00E03244: mov eax, [esi]
  loc_00E03246: call [eax+00000004h]
  loc_00E03249: mov cl, New_Opacity
  loc_00E0324C: mov edx, [esi]
  loc_00E0324E: push esi
  loc_00E0324F: mov [esi+00000159h], cl
  loc_00E03255: call [edx+00000910h]
  loc_00E0325B: sub esp, 00000010h
  loc_00E0325E: mov ecx, 00000008h
  loc_00E03263: mov edi, esp
  loc_00E03265: mov edx, [esi]
  loc_00E03267: mov eax, 006DED08h ; "PictureOpacityOnOver"
  loc_00E0326C: push esi
  loc_00E0326D: mov [edi], ecx
  loc_00E0326F: mov ecx, var_20
  loc_00E03272: mov [edi+00000004h], ecx
  loc_00E03275: mov [edi+00000008h], eax
  loc_00E03278: mov eax, var_18
  loc_00E0327B: mov [edi+0000000Ch], eax
  loc_00E0327E: call [edx+00000390h]
  loc_00E03284: test eax, eax
  loc_00E03286: fnclex
  loc_00E03288: jge 00E0329Ch
  loc_00E0328A: push 00000390h
  loc_00E0328F: push 006DCB00h
  loc_00E03294: push esi
  loc_00E03295: push eax
  loc_00E03296: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0329C: mov eax, Me
  loc_00E0329F: push eax
  loc_00E032A0: mov ecx, [eax]
  loc_00E032A2: call [ecx+00000008h]
  loc_00E032A5: mov eax, var_4
  loc_00E032A8: mov ecx, var_14
  loc_00E032AB: pop edi
  loc_00E032AC: pop esi
  loc_00E032AD: mov fs:[00000000h], ecx
  loc_00E032B4: pop ebx
  loc_00E032B5: mov esp, ebp
  loc_00E032B7: pop ebp
  loc_00E032B8: retn 0008h
End Sub

Public Property Get RightToLeft(arg_C) 'E032C0
  loc_00E032C0: push ebp
  loc_00E032C1: mov ebp, esp
  loc_00E032C3: sub esp, 0000000Ch
  loc_00E032C6: push 00402836h ; __vbaExceptHandler
  loc_00E032CB: mov eax, fs:[00000000h]
  loc_00E032D1: push eax
  loc_00E032D2: mov fs:[00000000h], esp
  loc_00E032D9: sub esp, 0000000Ch
  loc_00E032DC: push ebx
  loc_00E032DD: push esi
  loc_00E032DE: push edi
  loc_00E032DF: mov var_C, esp
  loc_00E032E2: mov var_8, 00401E40h
  loc_00E032E9: xor edi, edi
  loc_00E032EB: mov var_4, edi
  loc_00E032EE: mov esi, Me
  loc_00E032F1: push esi
  loc_00E032F2: mov eax, [esi]
  loc_00E032F4: call [eax+00000004h]
  loc_00E032F7: mov cx, [esi+00000138h]
  loc_00E032FE: mov var_18, edi
  loc_00E03301: mov var_18, ecx
  loc_00E03304: mov eax, Me
  loc_00E03307: push eax
  loc_00E03308: mov edx, [eax]
  loc_00E0330A: call [edx+00000008h]
  loc_00E0330D: mov eax, arg_C
  loc_00E03310: mov cx, var_18
  loc_00E03314: mov [eax], cx
  loc_00E03317: mov eax, var_4
  loc_00E0331A: mov ecx, var_14
  loc_00E0331D: pop edi
  loc_00E0331E: pop esi
  loc_00E0331F: mov fs:[00000000h], ecx
  loc_00E03326: pop ebx
  loc_00E03327: mov esp, ebp
  loc_00E03329: pop ebp
  loc_00E0332A: retn 0008h
End Sub

Public Property Let RightToLeft(Value) 'E03330
  loc_00E03330: push ebp
  loc_00E03331: mov ebp, esp
  loc_00E03333: sub esp, 0000000Ch
  loc_00E03336: push 00402836h ; __vbaExceptHandler
  loc_00E0333B: mov eax, fs:[00000000h]
  loc_00E03341: push eax
  loc_00E03342: mov fs:[00000000h], esp
  loc_00E03349: sub esp, 0000001Ch
  loc_00E0334C: push ebx
  loc_00E0334D: push esi
  loc_00E0334E: push edi
  loc_00E0334F: mov var_C, esp
  loc_00E03352: mov var_8, 00401E48h
  loc_00E03359: mov var_4, 00000000h
  loc_00E03360: mov esi, Me
  loc_00E03363: push esi
  loc_00E03364: mov eax, [esi]
  loc_00E03366: call [eax+00000004h]
  loc_00E03369: mov ax, Value
  loc_00E0336D: mov ecx, [esi]
  loc_00E0336F: push esi
  loc_00E03370: mov [esi+0000010Eh], ax
  loc_00E03377: mov [esi+00000138h], ax
  loc_00E0337E: call [ecx+00000910h]
  loc_00E03384: sub esp, 00000010h
  loc_00E03387: mov ecx, 00000008h
  loc_00E0338C: mov edi, esp
  loc_00E0338E: mov edx, [esi]
  loc_00E03390: mov eax, 006DEF8Ch ; "RightToLeft"
  loc_00E03395: push esi
  loc_00E03396: mov [edi], ecx
  loc_00E03398: mov ecx, var_20
  loc_00E0339B: mov [edi+00000004h], ecx
  loc_00E0339E: mov [edi+00000008h], eax
  loc_00E033A1: mov eax, var_18
  loc_00E033A4: mov [edi+0000000Ch], eax
  loc_00E033A7: call [edx+00000390h]
  loc_00E033AD: test eax, eax
  loc_00E033AF: fnclex
  loc_00E033B1: jge 00E033C5h
  loc_00E033B3: push 00000390h
  loc_00E033B8: push 006DCB00h
  loc_00E033BD: push esi
  loc_00E033BE: push eax
  loc_00E033BF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E033C5: mov eax, Me
  loc_00E033C8: push eax
  loc_00E033C9: mov ecx, [eax]
  loc_00E033CB: call [ecx+00000008h]
  loc_00E033CE: mov eax, var_4
  loc_00E033D1: mov ecx, var_14
  loc_00E033D4: pop edi
  loc_00E033D5: pop esi
  loc_00E033D6: mov fs:[00000000h], ecx
  loc_00E033DD: pop ebx
  loc_00E033DE: mov esp, ebp
  loc_00E033E0: pop ebp
  loc_00E033E1: retn 0008h
End Sub

Public Property Get CaptionEffects(arg_C) 'E033F0
  loc_00E033F0: push ebp
  loc_00E033F1: mov ebp, esp
  loc_00E033F3: sub esp, 0000000Ch
  loc_00E033F6: push 00402836h ; __vbaExceptHandler
  loc_00E033FB: mov eax, fs:[00000000h]
  loc_00E03401: push eax
  loc_00E03402: mov fs:[00000000h], esp
  loc_00E03409: sub esp, 0000000Ch
  loc_00E0340C: push ebx
  loc_00E0340D: push esi
  loc_00E0340E: push edi
  loc_00E0340F: mov var_C, esp
  loc_00E03412: mov var_8, 00401E50h
  loc_00E03419: xor edi, edi
  loc_00E0341B: mov var_4, edi
  loc_00E0341E: mov esi, Me
  loc_00E03421: push esi
  loc_00E03422: mov eax, [esi]
  loc_00E03424: call [eax+00000004h]
  loc_00E03427: mov ecx, [esi+00000068h]
  loc_00E0342A: mov var_18, edi
  loc_00E0342D: mov var_18, ecx
  loc_00E03430: mov eax, Me
  loc_00E03433: push eax
  loc_00E03434: mov edx, [eax]
  loc_00E03436: call [edx+00000008h]
  loc_00E03439: mov eax, arg_C
  loc_00E0343C: mov ecx, var_18
  loc_00E0343F: mov [eax], ecx
  loc_00E03441: mov eax, var_4
  loc_00E03444: mov ecx, var_14
  loc_00E03447: pop edi
  loc_00E03448: pop esi
  loc_00E03449: mov fs:[00000000h], ecx
  loc_00E03450: pop ebx
  loc_00E03451: mov esp, ebp
  loc_00E03453: pop ebp
  loc_00E03454: retn 0008h
End Sub

Public Property Let CaptionEffects(New_Effects) 'E03460
  loc_00E03460: push ebp
  loc_00E03461: mov ebp, esp
  loc_00E03463: sub esp, 0000000Ch
  loc_00E03466: push 00402836h ; __vbaExceptHandler
  loc_00E0346B: mov eax, fs:[00000000h]
  loc_00E03471: push eax
  loc_00E03472: mov fs:[00000000h], esp
  loc_00E03479: sub esp, 0000001Ch
  loc_00E0347C: push ebx
  loc_00E0347D: push esi
  loc_00E0347E: push edi
  loc_00E0347F: mov var_C, esp
  loc_00E03482: mov var_8, 00401E58h
  loc_00E03489: mov var_4, 00000000h
  loc_00E03490: mov esi, Me
  loc_00E03493: push esi
  loc_00E03494: mov eax, [esi]
  loc_00E03496: call [eax+00000004h]
  loc_00E03499: mov ecx, New_Effects
  loc_00E0349C: mov edx, [esi]
  loc_00E0349E: push esi
  loc_00E0349F: mov [esi+00000068h], ecx
  loc_00E034A2: call [edx+00000910h]
  loc_00E034A8: sub esp, 00000010h
  loc_00E034AB: mov ecx, 00000008h
  loc_00E034B0: mov edi, esp
  loc_00E034B2: mov edx, [esi]
  loc_00E034B4: mov eax, 006DEE20h ; "CaptionEffects"
  loc_00E034B9: push esi
  loc_00E034BA: mov [edi], ecx
  loc_00E034BC: mov ecx, var_20
  loc_00E034BF: mov [edi+00000004h], ecx
  loc_00E034C2: mov [edi+00000008h], eax
  loc_00E034C5: mov eax, var_18
  loc_00E034C8: mov [edi+0000000Ch], eax
  loc_00E034CB: call [edx+00000390h]
  loc_00E034D1: test eax, eax
  loc_00E034D3: fnclex
  loc_00E034D5: jge 00E034E9h
  loc_00E034D7: push 00000390h
  loc_00E034DC: push 006DCB00h
  loc_00E034E1: push esi
  loc_00E034E2: push eax
  loc_00E034E3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E034E9: mov eax, Me
  loc_00E034EC: push eax
  loc_00E034ED: mov ecx, [eax]
  loc_00E034EF: call [ecx+00000008h]
  loc_00E034F2: mov eax, var_4
  loc_00E034F5: mov ecx, var_14
  loc_00E034F8: pop edi
  loc_00E034F9: pop esi
  loc_00E034FA: mov fs:[00000000h], ecx
  loc_00E03501: pop ebx
  loc_00E03502: mov esp, ebp
  loc_00E03504: pop ebp
  loc_00E03505: retn 0008h
End Sub

Public Property Get ShowFocusRect(arg_C) 'E03510
  loc_00E03510: push ebp
  loc_00E03511: mov ebp, esp
  loc_00E03513: sub esp, 0000000Ch
  loc_00E03516: push 00402836h ; __vbaExceptHandler
  loc_00E0351B: mov eax, fs:[00000000h]
  loc_00E03521: push eax
  loc_00E03522: mov fs:[00000000h], esp
  loc_00E03529: sub esp, 0000000Ch
  loc_00E0352C: push ebx
  loc_00E0352D: push esi
  loc_00E0352E: push edi
  loc_00E0352F: mov var_C, esp
  loc_00E03532: mov var_8, 00401E60h
  loc_00E03539: xor edi, edi
  loc_00E0353B: mov var_4, edi
  loc_00E0353E: mov esi, Me
  loc_00E03541: push esi
  loc_00E03542: mov eax, [esi]
  loc_00E03544: call [eax+00000004h]
  loc_00E03547: mov cx, [esi+0000006Eh]
  loc_00E0354B: mov var_18, edi
  loc_00E0354E: mov var_18, ecx
  loc_00E03551: mov eax, Me
  loc_00E03554: push eax
  loc_00E03555: mov edx, [eax]
  loc_00E03557: call [edx+00000008h]
  loc_00E0355A: mov eax, arg_C
  loc_00E0355D: mov cx, var_18
  loc_00E03561: mov [eax], cx
  loc_00E03564: mov eax, var_4
  loc_00E03567: mov ecx, var_14
  loc_00E0356A: pop edi
  loc_00E0356B: pop esi
  loc_00E0356C: mov fs:[00000000h], ecx
  loc_00E03573: pop ebx
  loc_00E03574: mov esp, ebp
  loc_00E03576: pop ebp
  loc_00E03577: retn 0008h
End Sub

Public Property Let ShowFocusRect(New_ShowFocusRect) 'E03580
  loc_00E03580: push ebp
  loc_00E03581: mov ebp, esp
  loc_00E03583: sub esp, 0000000Ch
  loc_00E03586: push 00402836h ; __vbaExceptHandler
  loc_00E0358B: mov eax, fs:[00000000h]
  loc_00E03591: push eax
  loc_00E03592: mov fs:[00000000h], esp
  loc_00E03599: sub esp, 0000001Ch
  loc_00E0359C: push ebx
  loc_00E0359D: push esi
  loc_00E0359E: push edi
  loc_00E0359F: mov var_C, esp
  loc_00E035A2: mov var_8, 00401E68h
  loc_00E035A9: mov var_4, 00000000h
  loc_00E035B0: mov esi, Me
  loc_00E035B3: push esi
  loc_00E035B4: mov eax, [esi]
  loc_00E035B6: call [eax+00000004h]
  loc_00E035B9: mov cx, New_ShowFocusRect
  loc_00E035BD: sub esp, 00000010h
  loc_00E035C0: mov [esi+0000006Eh], cx
  loc_00E035C4: mov edi, esp
  loc_00E035C6: mov ecx, 00000008h
  loc_00E035CB: mov edx, [esi]
  loc_00E035CD: mov [edi], ecx
  loc_00E035CF: mov ecx, var_20
  loc_00E035D2: mov eax, 006DEB8Ch ; "ShowFocusRect"
  loc_00E035D7: push esi
  loc_00E035D8: mov [edi+00000004h], ecx
  loc_00E035DB: mov [edi+00000008h], eax
  loc_00E035DE: mov eax, var_18
  loc_00E035E1: mov [edi+0000000Ch], eax
  loc_00E035E4: call [edx+00000390h]
  loc_00E035EA: test eax, eax
  loc_00E035EC: fnclex
  loc_00E035EE: jge 00E03602h
  loc_00E035F0: push 00000390h
  loc_00E035F5: push 006DCB00h
  loc_00E035FA: push esi
  loc_00E035FB: push eax
  loc_00E035FC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E03602: mov eax, Me
  loc_00E03605: push eax
  loc_00E03606: mov ecx, [eax]
  loc_00E03608: call [ecx+00000008h]
  loc_00E0360B: mov eax, var_4
  loc_00E0360E: mov ecx, var_14
  loc_00E03611: pop edi
  loc_00E03612: pop esi
  loc_00E03613: mov fs:[00000000h], ecx
  loc_00E0361A: pop ebx
  loc_00E0361B: mov esp, ebp
  loc_00E0361D: pop ebp
  loc_00E0361E: retn 0008h
End Sub

Public Property Get UseMaskColor(arg_C) 'E03630
  loc_00E03630: push ebp
  loc_00E03631: mov ebp, esp
  loc_00E03633: sub esp, 0000000Ch
  loc_00E03636: push 00402836h ; __vbaExceptHandler
  loc_00E0363B: mov eax, fs:[00000000h]
  loc_00E03641: push eax
  loc_00E03642: mov fs:[00000000h], esp
  loc_00E03649: sub esp, 0000000Ch
  loc_00E0364C: push ebx
  loc_00E0364D: push esi
  loc_00E0364E: push edi
  loc_00E0364F: mov var_C, esp
  loc_00E03652: mov var_8, 00401E70h
  loc_00E03659: xor edi, edi
  loc_00E0365B: mov var_4, edi
  loc_00E0365E: mov esi, Me
  loc_00E03661: push esi
  loc_00E03662: mov eax, [esi]
  loc_00E03664: call [eax+00000004h]
  loc_00E03667: mov cx, [esi+0000009Ch]
  loc_00E0366E: mov var_18, edi
  loc_00E03671: mov var_18, ecx
  loc_00E03674: mov eax, Me
  loc_00E03677: push eax
  loc_00E03678: mov edx, [eax]
  loc_00E0367A: call [edx+00000008h]
  loc_00E0367D: mov eax, arg_C
  loc_00E03680: mov cx, var_18
  loc_00E03684: mov [eax], cx
  loc_00E03687: mov eax, var_4
  loc_00E0368A: mov ecx, var_14
  loc_00E0368D: pop edi
  loc_00E0368E: pop esi
  loc_00E0368F: mov fs:[00000000h], ecx
  loc_00E03696: pop ebx
  loc_00E03697: mov esp, ebp
  loc_00E03699: pop ebp
  loc_00E0369A: retn 0008h
End Sub

Public Property Let UseMaskColor(New_UseMaskColor) 'E036A0
  loc_00E036A0: push ebp
  loc_00E036A1: mov ebp, esp
  loc_00E036A3: sub esp, 0000000Ch
  loc_00E036A6: push 00402836h ; __vbaExceptHandler
  loc_00E036AB: mov eax, fs:[00000000h]
  loc_00E036B1: push eax
  loc_00E036B2: mov fs:[00000000h], esp
  loc_00E036B9: sub esp, 0000001Ch
  loc_00E036BC: push ebx
  loc_00E036BD: push esi
  loc_00E036BE: push edi
  loc_00E036BF: mov var_C, esp
  loc_00E036C2: mov var_8, 00401E78h
  loc_00E036C9: mov var_4, 00000000h
  loc_00E036D0: mov esi, Me
  loc_00E036D3: push esi
  loc_00E036D4: mov eax, [esi]
  loc_00E036D6: call [eax+00000004h]
  loc_00E036D9: mov edx, [esi+0000014Ch]
  loc_00E036DF: mov cx, New_UseMaskColor
  loc_00E036E3: push 00000000h
  loc_00E036E5: push edx
  loc_00E036E6: mov [esi+0000009Ch], cx
  loc_00E036ED: call [0040114Ch] ; __vbaObjIs
  loc_00E036F3: test ax, ax
  loc_00E036F6: jnz 00E03701h
  loc_00E036F8: mov eax, [esi]
  loc_00E036FA: push esi
  loc_00E036FB: call [eax+00000910h]
  loc_00E03701: sub esp, 00000010h
  loc_00E03704: mov ecx, 00000008h
  loc_00E03709: mov edi, esp
  loc_00E0370B: mov edx, [esi]
  loc_00E0370D: mov eax, 006DEE00h ; "UseMaskColor"
  loc_00E03712: push esi
  loc_00E03713: mov [edi], ecx
  loc_00E03715: mov ecx, var_20
  loc_00E03718: mov [edi+00000004h], ecx
  loc_00E0371B: mov [edi+00000008h], eax
  loc_00E0371E: mov eax, var_18
  loc_00E03721: mov [edi+0000000Ch], eax
  loc_00E03724: call [edx+00000390h]
  loc_00E0372A: test eax, eax
  loc_00E0372C: fnclex
  loc_00E0372E: jge 00E03742h
  loc_00E03730: push 00000390h
  loc_00E03735: push 006DCB00h
  loc_00E0373A: push esi
  loc_00E0373B: push eax
  loc_00E0373C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E03742: mov eax, Me
  loc_00E03745: push eax
  loc_00E03746: mov ecx, [eax]
  loc_00E03748: call [ecx+00000008h]
  loc_00E0374B: mov eax, var_4
  loc_00E0374E: mov ecx, var_14
  loc_00E03751: pop edi
  loc_00E03752: pop esi
  loc_00E03753: mov fs:[00000000h], ecx
  loc_00E0375A: pop ebx
  loc_00E0375B: mov esp, ebp
  loc_00E0375D: pop ebp
  loc_00E0375E: retn 0008h
End Sub

Public Property Get Value(arg_C) 'E03770
  loc_00E03770: push ebp
  loc_00E03771: mov ebp, esp
  loc_00E03773: sub esp, 0000000Ch
  loc_00E03776: push 00402836h ; __vbaExceptHandler
  loc_00E0377B: mov eax, fs:[00000000h]
  loc_00E03781: push eax
  loc_00E03782: mov fs:[00000000h], esp
  loc_00E03789: sub esp, 0000000Ch
  loc_00E0378C: push ebx
  loc_00E0378D: push esi
  loc_00E0378E: push edi
  loc_00E0378F: mov var_C, esp
  loc_00E03792: mov var_8, 00401E80h
  loc_00E03799: xor edi, edi
  loc_00E0379B: mov var_4, edi
  loc_00E0379E: mov esi, Me
  loc_00E037A1: push esi
  loc_00E037A2: mov eax, [esi]
  loc_00E037A4: call [eax+00000004h]
  loc_00E037A7: mov cx, [esi+0000006Ch]
  loc_00E037AB: mov var_18, edi
  loc_00E037AE: mov var_18, ecx
  loc_00E037B1: mov eax, Me
  loc_00E037B4: push eax
  loc_00E037B5: mov edx, [eax]
  loc_00E037B7: call [edx+00000008h]
  loc_00E037BA: mov eax, arg_C
  loc_00E037BD: mov cx, var_18
  loc_00E037C1: mov [eax], cx
  loc_00E037C4: mov eax, var_4
  loc_00E037C7: mov ecx, var_14
  loc_00E037CA: pop edi
  loc_00E037CB: pop esi
  loc_00E037CC: mov fs:[00000000h], ecx
  loc_00E037D3: pop ebx
  loc_00E037D4: mov esp, ebp
  loc_00E037D6: pop ebp
  loc_00E037D7: retn 0008h
End Sub

Public Property Let Value(New_Value) 'E037E0
  loc_00E037E0: push ebp
  loc_00E037E1: mov ebp, esp
  loc_00E037E3: sub esp, 0000000Ch
  loc_00E037E6: push 00402836h ; __vbaExceptHandler
  loc_00E037EB: mov eax, fs:[00000000h]
  loc_00E037F1: push eax
  loc_00E037F2: mov fs:[00000000h], esp
  loc_00E037F9: sub esp, 0000001Ch
  loc_00E037FC: push ebx
  loc_00E037FD: push esi
  loc_00E037FE: push edi
  loc_00E037FF: mov var_C, esp
  loc_00E03802: mov var_8, 00401E88h
  loc_00E03809: xor edi, edi
  loc_00E0380B: mov var_4, edi
  loc_00E0380E: mov esi, Me
  loc_00E03811: push esi
  loc_00E03812: mov eax, [esi]
  loc_00E03814: call [eax+00000004h]
  loc_00E03817: mov eax, [esi+00000064h]
  loc_00E0381A: mov var_24, edi
  loc_00E0381D: cmp eax, edi
  loc_00E0381F: jz 00E0387Dh
  loc_00E03821: mov ax, New_Value
  loc_00E03825: cmp ax, di
  loc_00E03828: mov [esi+0000006Ch], ax
  loc_00E0382C: jnz 00E03831h
  loc_00E0382E: mov [esi+00000048h], edi
  loc_00E03831: mov ecx, [esi]
  loc_00E03833: push esi
  loc_00E03834: call [ecx+00000910h]
  loc_00E0383A: sub esp, 00000010h
  loc_00E0383D: mov ecx, 00000008h
  loc_00E03842: mov ebx, esp
  loc_00E03844: mov edx, [esi]
  loc_00E03846: mov eax, 006DEBFCh ; "Value"
  loc_00E0384B: push esi
  loc_00E0384C: mov [ebx], ecx
  loc_00E0384E: mov ecx, var_20
  loc_00E03851: mov [ebx+00000004h], ecx
  loc_00E03854: mov [ebx+00000008h], eax
  loc_00E03857: mov eax, var_18
  loc_00E0385A: mov [ebx+0000000Ch], eax
  loc_00E0385D: call [edx+00000390h]
  loc_00E03863: cmp eax, edi
  loc_00E03865: fnclex
  loc_00E03867: jge 00E03889h
  loc_00E03869: push 00000390h
  loc_00E0386E: push 006DCB00h
  loc_00E03873: push esi
  loc_00E03874: push eax
  loc_00E03875: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0387B: jmp 00E03889h
  loc_00E0387D: mov ecx, [esi]
  loc_00E0387F: push esi
  loc_00E03880: mov [esi+00000048h], edi
  loc_00E03883: call [ecx+00000910h]
  loc_00E03889: mov eax, Me
  loc_00E0388C: push eax
  loc_00E0388D: mov edx, [eax]
  loc_00E0388F: call [edx+00000008h]
  loc_00E03892: mov eax, var_4
  loc_00E03895: mov ecx, var_14
  loc_00E03898: pop edi
  loc_00E03899: pop esi
  loc_00E0389A: mov fs:[00000000h], ecx
  loc_00E038A1: pop ebx
  loc_00E038A2: mov esp, ebp
  loc_00E038A4: pop ebp
  loc_00E038A5: retn 0008h
End Sub

Public Property Get TooltipTitle(arg_C) 'E038B0
  loc_00E038B0: push ebp
  loc_00E038B1: mov ebp, esp
  loc_00E038B3: sub esp, 0000000Ch
  loc_00E038B6: push 00402836h ; __vbaExceptHandler
  loc_00E038BB: mov eax, fs:[00000000h]
  loc_00E038C1: push eax
  loc_00E038C2: mov fs:[00000000h], esp
  loc_00E038C9: sub esp, 0000000Ch
  loc_00E038CC: push ebx
  loc_00E038CD: push esi
  loc_00E038CE: push edi
  loc_00E038CF: mov var_C, esp
  loc_00E038D2: mov var_8, 00401E90h
  loc_00E038D9: xor edi, edi
  loc_00E038DB: mov var_4, edi
  loc_00E038DE: mov esi, Me
  loc_00E038E1: push esi
  loc_00E038E2: mov eax, [esi]
  loc_00E038E4: call [eax+00000004h]
  loc_00E038E7: mov ecx, arg_C
  loc_00E038EA: mov var_18, edi
  loc_00E038ED: mov [ecx], edi
  loc_00E038EF: mov edx, [esi+000000FCh]
  loc_00E038F5: lea ecx, var_18
  loc_00E038F8: call [004011B0h] ; __vbaStrCopy
  loc_00E038FE: push 00E03910h
  loc_00E03903: jmp 00E0390Fh
  loc_00E03905: lea ecx, var_18
  loc_00E03908: call [00401258h] ; __vbaFreeStr
  loc_00E0390E: ret
  loc_00E0390F: ret
  loc_00E03910: mov eax, Me
  loc_00E03913: push eax
  loc_00E03914: mov edx, [eax]
  loc_00E03916: call [edx+00000008h]
  loc_00E03919: mov eax, arg_C
  loc_00E0391C: mov ecx, var_18
  loc_00E0391F: mov [eax], ecx
  loc_00E03921: mov eax, var_4
  loc_00E03924: mov ecx, var_14
  loc_00E03927: pop edi
  loc_00E03928: pop esi
  loc_00E03929: mov fs:[00000000h], ecx
  loc_00E03930: pop ebx
  loc_00E03931: mov esp, ebp
  loc_00E03933: pop ebp
  loc_00E03934: retn 0008h
End Sub

Public Property Let TooltipTitle(New_title) 'E03940
  loc_00E03940: push ebp
  loc_00E03941: mov ebp, esp
  loc_00E03943: sub esp, 0000000Ch
  loc_00E03946: push 00402836h ; __vbaExceptHandler
  loc_00E0394B: mov eax, fs:[00000000h]
  loc_00E03951: push eax
  loc_00E03952: mov fs:[00000000h], esp
  loc_00E03959: sub esp, 00000020h
  loc_00E0395C: push ebx
  loc_00E0395D: push esi
  loc_00E0395E: push edi
  loc_00E0395F: mov var_C, esp
  loc_00E03962: mov var_8, 00401EA0h
  loc_00E03969: xor ebx, ebx
  loc_00E0396B: mov var_4, ebx
  loc_00E0396E: mov esi, Me
  loc_00E03971: push esi
  loc_00E03972: mov eax, [esi]
  loc_00E03974: call [eax+00000004h]
  loc_00E03977: mov edx, New_title
  loc_00E0397A: mov edi, [004011B0h] ; __vbaStrCopy
  loc_00E03980: lea ecx, var_18
  loc_00E03983: mov var_18, ebx
  loc_00E03986: call edi
  loc_00E03988: mov edx, var_18
  loc_00E0398B: lea ecx, [esi+000000FCh]
  loc_00E03991: call edi
  loc_00E03993: mov ecx, [esi]
  loc_00E03995: push esi
  loc_00E03996: call [ecx+00000910h]
  loc_00E0399C: sub esp, 00000010h
  loc_00E0399F: mov ecx, 00000008h
  loc_00E039A4: mov edi, esp
  loc_00E039A6: mov edx, [esi]
  loc_00E039A8: mov eax, 006DEEF8h ; "TooltipTitle"
  loc_00E039AD: push esi
  loc_00E039AE: mov [edi], ecx
  loc_00E039B0: mov ecx, var_24
  loc_00E039B3: mov [edi+00000004h], ecx
  loc_00E039B6: mov [edi+00000008h], eax
  loc_00E039B9: mov eax, var_1C
  loc_00E039BC: mov [edi+0000000Ch], eax
  loc_00E039BF: call [edx+00000390h]
  loc_00E039C5: cmp eax, ebx
  loc_00E039C7: fnclex
  loc_00E039C9: jge 00E039DDh
  loc_00E039CB: push 00000390h
  loc_00E039D0: push 006DCB00h
  loc_00E039D5: push esi
  loc_00E039D6: push eax
  loc_00E039D7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E039DD: push 00E039ECh
  loc_00E039E2: lea ecx, var_18
  loc_00E039E5: call [00401258h] ; __vbaFreeStr
  loc_00E039EB: ret
  loc_00E039EC: mov eax, Me
  loc_00E039EF: push eax
  loc_00E039F0: mov ecx, [eax]
  loc_00E039F2: call [ecx+00000008h]
  loc_00E039F5: mov eax, var_4
  loc_00E039F8: mov ecx, var_14
  loc_00E039FB: pop edi
  loc_00E039FC: pop esi
  loc_00E039FD: mov fs:[00000000h], ecx
  loc_00E03A04: pop ebx
  loc_00E03A05: mov esp, ebp
  loc_00E03A07: pop ebp
  loc_00E03A08: retn 0008h
End Sub

Public Property Get ToolTip(arg_C) 'E03A10
  loc_00E03A10: push ebp
  loc_00E03A11: mov ebp, esp
  loc_00E03A13: sub esp, 0000000Ch
  loc_00E03A16: push 00402836h ; __vbaExceptHandler
  loc_00E03A1B: mov eax, fs:[00000000h]
  loc_00E03A21: push eax
  loc_00E03A22: mov fs:[00000000h], esp
  loc_00E03A29: sub esp, 0000000Ch
  loc_00E03A2C: push ebx
  loc_00E03A2D: push esi
  loc_00E03A2E: push edi
  loc_00E03A2F: mov var_C, esp
  loc_00E03A32: mov var_8, 00401EB0h
  loc_00E03A39: xor edi, edi
  loc_00E03A3B: mov var_4, edi
  loc_00E03A3E: mov esi, Me
  loc_00E03A41: push esi
  loc_00E03A42: mov eax, [esi]
  loc_00E03A44: call [eax+00000004h]
  loc_00E03A47: mov ecx, arg_C
  loc_00E03A4A: mov var_18, edi
  loc_00E03A4D: mov [ecx], edi
  loc_00E03A4F: mov edx, [esi+000000F8h]
  loc_00E03A55: lea ecx, var_18
  loc_00E03A58: call [004011B0h] ; __vbaStrCopy
  loc_00E03A5E: push 00E03A70h
  loc_00E03A63: jmp 00E03A6Fh
  loc_00E03A65: lea ecx, var_18
  loc_00E03A68: call [00401258h] ; __vbaFreeStr
  loc_00E03A6E: ret
  loc_00E03A6F: ret
  loc_00E03A70: mov eax, Me
  loc_00E03A73: push eax
  loc_00E03A74: mov edx, [eax]
  loc_00E03A76: call [edx+00000008h]
  loc_00E03A79: mov eax, arg_C
  loc_00E03A7C: mov ecx, var_18
  loc_00E03A7F: mov [eax], ecx
  loc_00E03A81: mov eax, var_4
  loc_00E03A84: mov ecx, var_14
  loc_00E03A87: pop edi
  loc_00E03A88: pop esi
  loc_00E03A89: mov fs:[00000000h], ecx
  loc_00E03A90: pop ebx
  loc_00E03A91: mov esp, ebp
  loc_00E03A93: pop ebp
  loc_00E03A94: retn 0008h
End Sub

Public Property Let ToolTip(New_Tooltip) 'E03AA0
  loc_00E03AA0: push ebp
  loc_00E03AA1: mov ebp, esp
  loc_00E03AA3: sub esp, 0000000Ch
  loc_00E03AA6: push 00402836h ; __vbaExceptHandler
  loc_00E03AAB: mov eax, fs:[00000000h]
  loc_00E03AB1: push eax
  loc_00E03AB2: mov fs:[00000000h], esp
  loc_00E03AB9: sub esp, 00000020h
  loc_00E03ABC: push ebx
  loc_00E03ABD: push esi
  loc_00E03ABE: push edi
  loc_00E03ABF: mov var_C, esp
  loc_00E03AC2: mov var_8, 00401EC0h
  loc_00E03AC9: xor ebx, ebx
  loc_00E03ACB: mov var_4, ebx
  loc_00E03ACE: mov esi, Me
  loc_00E03AD1: push esi
  loc_00E03AD2: mov eax, [esi]
  loc_00E03AD4: call [eax+00000004h]
  loc_00E03AD7: mov edx, New_Tooltip
  loc_00E03ADA: mov edi, [004011B0h] ; __vbaStrCopy
  loc_00E03AE0: lea ecx, var_18
  loc_00E03AE3: mov var_18, ebx
  loc_00E03AE6: call edi
  loc_00E03AE8: mov edx, var_18
  loc_00E03AEB: lea ecx, [esi+000000F8h]
  loc_00E03AF1: call edi
  loc_00E03AF3: mov ecx, [esi]
  loc_00E03AF5: push esi
  loc_00E03AF6: call [ecx+00000910h]
  loc_00E03AFC: sub esp, 00000010h
  loc_00E03AFF: mov ecx, 00000008h
  loc_00E03B04: mov edi, esp
  loc_00E03B06: mov edx, [esi]
  loc_00E03B08: mov eax, 006DEF18h ; "ToolTip"
  loc_00E03B0D: push esi
  loc_00E03B0E: mov [edi], ecx
  loc_00E03B10: mov ecx, var_24
  loc_00E03B13: mov [edi+00000004h], ecx
  loc_00E03B16: mov [edi+00000008h], eax
  loc_00E03B19: mov eax, var_1C
  loc_00E03B1C: mov [edi+0000000Ch], eax
  loc_00E03B1F: call [edx+00000390h]
  loc_00E03B25: cmp eax, ebx
  loc_00E03B27: fnclex
  loc_00E03B29: jge 00E03B3Dh
  loc_00E03B2B: push 00000390h
  loc_00E03B30: push 006DCB00h
  loc_00E03B35: push esi
  loc_00E03B36: push eax
  loc_00E03B37: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E03B3D: push 00E03B4Ch
  loc_00E03B42: lea ecx, var_18
  loc_00E03B45: call [00401258h] ; __vbaFreeStr
  loc_00E03B4B: ret
  loc_00E03B4C: mov eax, Me
  loc_00E03B4F: push eax
  loc_00E03B50: mov ecx, [eax]
  loc_00E03B52: call [ecx+00000008h]
  loc_00E03B55: mov eax, var_4
  loc_00E03B58: mov ecx, var_14
  loc_00E03B5B: pop edi
  loc_00E03B5C: pop esi
  loc_00E03B5D: mov fs:[00000000h], ecx
  loc_00E03B64: pop ebx
  loc_00E03B65: mov esp, ebp
  loc_00E03B67: pop ebp
  loc_00E03B68: retn 0008h
End Sub

Public Property Get TooltipBackColor(arg_C) 'E03B70
  loc_00E03B70: push ebp
  loc_00E03B71: mov ebp, esp
  loc_00E03B73: sub esp, 0000000Ch
  loc_00E03B76: push 00402836h ; __vbaExceptHandler
  loc_00E03B7B: mov eax, fs:[00000000h]
  loc_00E03B81: push eax
  loc_00E03B82: mov fs:[00000000h], esp
  loc_00E03B89: sub esp, 0000000Ch
  loc_00E03B8C: push ebx
  loc_00E03B8D: push esi
  loc_00E03B8E: push edi
  loc_00E03B8F: mov var_C, esp
  loc_00E03B92: mov var_8, 00401ED0h
  loc_00E03B99: xor edi, edi
  loc_00E03B9B: mov var_4, edi
  loc_00E03B9E: mov esi, Me
  loc_00E03BA1: push esi
  loc_00E03BA2: mov eax, [esi]
  loc_00E03BA4: call [eax+00000004h]
  loc_00E03BA7: mov ecx, [esi+00000108h]
  loc_00E03BAD: mov var_18, edi
  loc_00E03BB0: mov var_18, ecx
  loc_00E03BB3: mov eax, Me
  loc_00E03BB6: push eax
  loc_00E03BB7: mov edx, [eax]
  loc_00E03BB9: call [edx+00000008h]
  loc_00E03BBC: mov eax, arg_C
  loc_00E03BBF: mov ecx, var_18
  loc_00E03BC2: mov [eax], ecx
  loc_00E03BC4: mov eax, var_4
  loc_00E03BC7: mov ecx, var_14
  loc_00E03BCA: pop edi
  loc_00E03BCB: pop esi
  loc_00E03BCC: mov fs:[00000000h], ecx
  loc_00E03BD3: pop ebx
  loc_00E03BD4: mov esp, ebp
  loc_00E03BD6: pop ebp
  loc_00E03BD7: retn 0008h
End Sub

Public Property Let TooltipBackColor(New_Color) 'E03BE0
  loc_00E03BE0: push ebp
  loc_00E03BE1: mov ebp, esp
  loc_00E03BE3: sub esp, 0000000Ch
  loc_00E03BE6: push 00402836h ; __vbaExceptHandler
  loc_00E03BEB: mov eax, fs:[00000000h]
  loc_00E03BF1: push eax
  loc_00E03BF2: mov fs:[00000000h], esp
  loc_00E03BF9: sub esp, 0000001Ch
  loc_00E03BFC: push ebx
  loc_00E03BFD: push esi
  loc_00E03BFE: push edi
  loc_00E03BFF: mov var_C, esp
  loc_00E03C02: mov var_8, 00401ED8h
  loc_00E03C09: mov var_4, 00000000h
  loc_00E03C10: mov esi, Me
  loc_00E03C13: push esi
  loc_00E03C14: mov eax, [esi]
  loc_00E03C16: call [eax+00000004h]
  loc_00E03C19: mov ecx, New_Color
  loc_00E03C1C: mov edx, [esi]
  loc_00E03C1E: push esi
  loc_00E03C1F: mov [esi+00000108h], ecx
  loc_00E03C25: call [edx+00000910h]
  loc_00E03C2B: sub esp, 00000010h
  loc_00E03C2E: mov ecx, 00000008h
  loc_00E03C33: mov edi, esp
  loc_00E03C35: mov edx, [esi]
  loc_00E03C37: mov eax, 006DF228h ; "TooltipBackcolor"
  loc_00E03C3C: push esi
  loc_00E03C3D: mov [edi], ecx
  loc_00E03C3F: mov ecx, var_20
  loc_00E03C42: mov [edi+00000004h], ecx
  loc_00E03C45: mov [edi+00000008h], eax
  loc_00E03C48: mov eax, var_18
  loc_00E03C4B: mov [edi+0000000Ch], eax
  loc_00E03C4E: call [edx+00000390h]
  loc_00E03C54: test eax, eax
  loc_00E03C56: fnclex
  loc_00E03C58: jge 00E03C6Ch
  loc_00E03C5A: push 00000390h
  loc_00E03C5F: push 006DCB00h
  loc_00E03C64: push esi
  loc_00E03C65: push eax
  loc_00E03C66: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E03C6C: mov eax, Me
  loc_00E03C6F: push eax
  loc_00E03C70: mov ecx, [eax]
  loc_00E03C72: call [ecx+00000008h]
  loc_00E03C75: mov eax, var_4
  loc_00E03C78: mov ecx, var_14
  loc_00E03C7B: pop edi
  loc_00E03C7C: pop esi
  loc_00E03C7D: mov fs:[00000000h], ecx
  loc_00E03C84: pop ebx
  loc_00E03C85: mov esp, ebp
  loc_00E03C87: pop ebp
  loc_00E03C88: retn 0008h
End Sub

Public Property Let ToolTipIcon(lTooltipIcon) 'E03C90
  loc_00E03C90: push ebp
  loc_00E03C91: mov ebp, esp
  loc_00E03C93: sub esp, 0000000Ch
  loc_00E03C96: push 00402836h ; __vbaExceptHandler
  loc_00E03C9B: mov eax, fs:[00000000h]
  loc_00E03CA1: push eax
  loc_00E03CA2: mov fs:[00000000h], esp
  loc_00E03CA9: sub esp, 0000001Ch
  loc_00E03CAC: push ebx
  loc_00E03CAD: push esi
  loc_00E03CAE: push edi
  loc_00E03CAF: mov var_C, esp
  loc_00E03CB2: mov var_8, 00401EE0h
  loc_00E03CB9: mov var_4, 00000000h
  loc_00E03CC0: mov esi, Me
  loc_00E03CC3: push esi
  loc_00E03CC4: mov eax, [esi]
  loc_00E03CC6: call [eax+00000004h]
  loc_00E03CC9: mov ecx, lTooltipIcon
  loc_00E03CCC: mov eax, [esi]
  loc_00E03CCE: push esi
  loc_00E03CCF: mov edx, [ecx]
  loc_00E03CD1: mov [esi+00000100h], edx
  loc_00E03CD7: call [eax+00000910h]
  loc_00E03CDD: sub esp, 00000010h
  loc_00E03CE0: mov ecx, 00000008h
  loc_00E03CE5: mov edi, esp
  loc_00E03CE7: mov edx, [esi]
  loc_00E03CE9: mov eax, 006DEF2Ch ; "TooltipIcon"
  loc_00E03CEE: push esi
  loc_00E03CEF: mov [edi], ecx
  loc_00E03CF1: mov ecx, var_20
  loc_00E03CF4: mov [edi+00000004h], ecx
  loc_00E03CF7: mov [edi+00000008h], eax
  loc_00E03CFA: mov eax, var_18
  loc_00E03CFD: mov [edi+0000000Ch], eax
  loc_00E03D00: call [edx+00000390h]
  loc_00E03D06: test eax, eax
  loc_00E03D08: fnclex
  loc_00E03D0A: jge 00E03D1Eh
  loc_00E03D0C: push 00000390h
  loc_00E03D11: push 006DCB00h
  loc_00E03D16: push esi
  loc_00E03D17: push eax
  loc_00E03D18: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E03D1E: mov eax, Me
  loc_00E03D21: push eax
  loc_00E03D22: mov ecx, [eax]
  loc_00E03D24: call [ecx+00000008h]
  loc_00E03D27: mov eax, var_4
  loc_00E03D2A: mov ecx, var_14
  loc_00E03D2D: pop edi
  loc_00E03D2E: pop esi
  loc_00E03D2F: mov fs:[00000000h], ecx
  loc_00E03D36: pop ebx
  loc_00E03D37: mov esp, ebp
  loc_00E03D39: pop ebp
  loc_00E03D3A: retn 0008h
End Sub

Public Property Get ToolTipIcon(arg_C) 'E03D40
  loc_00E03D40: push ebp
  loc_00E03D41: mov ebp, esp
  loc_00E03D43: sub esp, 0000000Ch
  loc_00E03D46: push 00402836h ; __vbaExceptHandler
  loc_00E03D4B: mov eax, fs:[00000000h]
  loc_00E03D51: push eax
  loc_00E03D52: mov fs:[00000000h], esp
  loc_00E03D59: sub esp, 0000000Ch
  loc_00E03D5C: push ebx
  loc_00E03D5D: push esi
  loc_00E03D5E: push edi
  loc_00E03D5F: mov var_C, esp
  loc_00E03D62: mov var_8, 00401EE8h
  loc_00E03D69: xor edi, edi
  loc_00E03D6B: mov var_4, edi
  loc_00E03D6E: mov esi, Me
  loc_00E03D71: push esi
  loc_00E03D72: mov eax, [esi]
  loc_00E03D74: call [eax+00000004h]
  loc_00E03D77: mov ecx, [esi+00000100h]
  loc_00E03D7D: mov var_18, edi
  loc_00E03D80: mov var_18, ecx
  loc_00E03D83: mov eax, Me
  loc_00E03D86: push eax
  loc_00E03D87: mov edx, [eax]
  loc_00E03D89: call [edx+00000008h]
  loc_00E03D8C: mov eax, arg_C
  loc_00E03D8F: mov ecx, var_18
  loc_00E03D92: mov [eax], ecx
  loc_00E03D94: mov eax, var_4
  loc_00E03D97: mov ecx, var_14
  loc_00E03D9A: pop edi
  loc_00E03D9B: pop esi
  loc_00E03D9C: mov fs:[00000000h], ecx
  loc_00E03DA3: pop ebx
  loc_00E03DA4: mov esp, ebp
  loc_00E03DA6: pop ebp
  loc_00E03DA7: retn 0008h
End Sub

Public Property Get ToolTipType(arg_C) 'E03DB0
  loc_00E03DB0: push ebp
  loc_00E03DB1: mov ebp, esp
  loc_00E03DB3: sub esp, 0000000Ch
  loc_00E03DB6: push 00402836h ; __vbaExceptHandler
  loc_00E03DBB: mov eax, fs:[00000000h]
  loc_00E03DC1: push eax
  loc_00E03DC2: mov fs:[00000000h], esp
  loc_00E03DC9: sub esp, 0000000Ch
  loc_00E03DCC: push ebx
  loc_00E03DCD: push esi
  loc_00E03DCE: push edi
  loc_00E03DCF: mov var_C, esp
  loc_00E03DD2: mov var_8, 00401EF0h
  loc_00E03DD9: xor edi, edi
  loc_00E03DDB: mov var_4, edi
  loc_00E03DDE: mov esi, Me
  loc_00E03DE1: push esi
  loc_00E03DE2: mov eax, [esi]
  loc_00E03DE4: call [eax+00000004h]
  loc_00E03DE7: mov ecx, [esi+00000104h]
  loc_00E03DED: mov var_18, edi
  loc_00E03DF0: mov var_18, ecx
  loc_00E03DF3: mov eax, Me
  loc_00E03DF6: push eax
  loc_00E03DF7: mov edx, [eax]
  loc_00E03DF9: call [edx+00000008h]
  loc_00E03DFC: mov eax, arg_C
  loc_00E03DFF: mov ecx, var_18
  loc_00E03E02: mov [eax], ecx
  loc_00E03E04: mov eax, var_4
  loc_00E03E07: mov ecx, var_14
  loc_00E03E0A: pop edi
  loc_00E03E0B: pop esi
  loc_00E03E0C: mov fs:[00000000h], ecx
  loc_00E03E13: pop ebx
  loc_00E03E14: mov esp, ebp
  loc_00E03E16: pop ebp
  loc_00E03E17: retn 0008h
End Sub

Public Property Let ToolTipType(lNewTTType) 'E03E20
  loc_00E03E20: push ebp
  loc_00E03E21: mov ebp, esp
  loc_00E03E23: sub esp, 0000000Ch
  loc_00E03E26: push 00402836h ; __vbaExceptHandler
  loc_00E03E2B: mov eax, fs:[00000000h]
  loc_00E03E31: push eax
  loc_00E03E32: mov fs:[00000000h], esp
  loc_00E03E39: sub esp, 0000001Ch
  loc_00E03E3C: push ebx
  loc_00E03E3D: push esi
  loc_00E03E3E: push edi
  loc_00E03E3F: mov var_C, esp
  loc_00E03E42: mov var_8, 00401EF8h
  loc_00E03E49: mov var_4, 00000000h
  loc_00E03E50: mov esi, Me
  loc_00E03E53: push esi
  loc_00E03E54: mov eax, [esi]
  loc_00E03E56: call [eax+00000004h]
  loc_00E03E59: mov ecx, lNewTTType
  loc_00E03E5C: mov edx, [esi]
  loc_00E03E5E: push esi
  loc_00E03E5F: mov [esi+00000104h], ecx
  loc_00E03E65: call [edx+00000910h]
  loc_00E03E6B: sub esp, 00000010h
  loc_00E03E6E: mov ecx, 00000008h
  loc_00E03E73: mov edi, esp
  loc_00E03E75: mov edx, [esi]
  loc_00E03E77: mov eax, 006DF250h ; "ToolTipType"
  loc_00E03E7C: push esi
  loc_00E03E7D: mov [edi], ecx
  loc_00E03E7F: mov ecx, var_20
  loc_00E03E82: mov [edi+00000004h], ecx
  loc_00E03E85: mov [edi+00000008h], eax
  loc_00E03E88: mov eax, var_18
  loc_00E03E8B: mov [edi+0000000Ch], eax
  loc_00E03E8E: call [edx+00000390h]
  loc_00E03E94: test eax, eax
  loc_00E03E96: fnclex
  loc_00E03E98: jge 00E03EACh
  loc_00E03E9A: push 00000390h
  loc_00E03E9F: push 006DCB00h
  loc_00E03EA4: push esi
  loc_00E03EA5: push eax
  loc_00E03EA6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E03EAC: mov eax, Me
  loc_00E03EAF: push eax
  loc_00E03EB0: mov ecx, [eax]
  loc_00E03EB2: call [ecx+00000008h]
  loc_00E03EB5: mov eax, var_4
  loc_00E03EB8: mov ecx, var_14
  loc_00E03EBB: pop edi
  loc_00E03EBC: pop esi
  loc_00E03EBD: mov fs:[00000000h], ecx
  loc_00E03EC4: pop ebx
  loc_00E03EC5: mov esp, ebp
  loc_00E03EC7: pop ebp
  loc_00E03EC8: retn 0008h
End Sub

Public Property Get ColorScheme(arg_C) 'E03ED0
  loc_00E03ED0: push ebp
  loc_00E03ED1: mov ebp, esp
  loc_00E03ED3: sub esp, 0000000Ch
  loc_00E03ED6: push 00402836h ; __vbaExceptHandler
  loc_00E03EDB: mov eax, fs:[00000000h]
  loc_00E03EE1: push eax
  loc_00E03EE2: mov fs:[00000000h], esp
  loc_00E03EE9: sub esp, 0000000Ch
  loc_00E03EEC: push ebx
  loc_00E03EED: push esi
  loc_00E03EEE: push edi
  loc_00E03EEF: mov var_C, esp
  loc_00E03EF2: mov var_8, 00401F00h
  loc_00E03EF9: xor edi, edi
  loc_00E03EFB: mov var_4, edi
  loc_00E03EFE: mov esi, Me
  loc_00E03F01: push esi
  loc_00E03F02: mov eax, [esi]
  loc_00E03F04: call [eax+00000004h]
  loc_00E03F07: mov ecx, [esi+000000CCh]
  loc_00E03F0D: mov var_18, edi
  loc_00E03F10: mov var_18, ecx
  loc_00E03F13: mov eax, Me
  loc_00E03F16: push eax
  loc_00E03F17: mov edx, [eax]
  loc_00E03F19: call [edx+00000008h]
  loc_00E03F1C: mov eax, arg_C
  loc_00E03F1F: mov ecx, var_18
  loc_00E03F22: mov [eax], ecx
  loc_00E03F24: mov eax, var_4
  loc_00E03F27: mov ecx, var_14
  loc_00E03F2A: pop edi
  loc_00E03F2B: pop esi
  loc_00E03F2C: mov fs:[00000000h], ecx
  loc_00E03F33: pop ebx
  loc_00E03F34: mov esp, ebp
  loc_00E03F36: pop ebp
  loc_00E03F37: retn 0008h
End Sub

Public Property Let ColorScheme(New_Color) 'E03F40
  loc_00E03F40: push ebp
  loc_00E03F41: mov ebp, esp
  loc_00E03F43: sub esp, 0000000Ch
  loc_00E03F46: push 00402836h ; __vbaExceptHandler
  loc_00E03F4B: mov eax, fs:[00000000h]
  loc_00E03F51: push eax
  loc_00E03F52: mov fs:[00000000h], esp
  loc_00E03F59: sub esp, 0000001Ch
  loc_00E03F5C: push ebx
  loc_00E03F5D: push esi
  loc_00E03F5E: push edi
  loc_00E03F5F: mov var_C, esp
  loc_00E03F62: mov var_8, 00401F08h
  loc_00E03F69: mov var_4, 00000000h
  loc_00E03F70: mov esi, Me
  loc_00E03F73: push esi
  loc_00E03F74: mov eax, [esi]
  loc_00E03F76: call [eax+00000004h]
  loc_00E03F79: mov ecx, New_Color
  loc_00E03F7C: mov edx, [esi]
  loc_00E03F7E: push esi
  loc_00E03F7F: mov [esi+000000CCh], ecx
  loc_00E03F85: call [edx+000009B8h]
  loc_00E03F8B: mov eax, [esi]
  loc_00E03F8D: push esi
  loc_00E03F8E: call [eax+00000910h]
  loc_00E03F94: sub esp, 00000010h
  loc_00E03F97: mov ecx, 00000008h
  loc_00E03F9C: mov edi, esp
  loc_00E03F9E: mov edx, [esi]
  loc_00E03FA0: mov eax, 006DEFE0h ; "ColorScheme"
  loc_00E03FA5: push esi
  loc_00E03FA6: mov [edi], ecx
  loc_00E03FA8: mov ecx, var_20
  loc_00E03FAB: mov [edi+00000004h], ecx
  loc_00E03FAE: mov [edi+00000008h], eax
  loc_00E03FB1: mov eax, var_18
  loc_00E03FB4: mov [edi+0000000Ch], eax
  loc_00E03FB7: call [edx+00000390h]
  loc_00E03FBD: test eax, eax
  loc_00E03FBF: fnclex
  loc_00E03FC1: jge 00E03FD5h
  loc_00E03FC3: push 00000390h
  loc_00E03FC8: push 006DCB00h
  loc_00E03FCD: push esi
  loc_00E03FCE: push eax
  loc_00E03FCF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E03FD5: mov eax, Me
  loc_00E03FD8: push eax
  loc_00E03FD9: mov ecx, [eax]
  loc_00E03FDB: call [ecx+00000008h]
  loc_00E03FDE: mov eax, var_4
  loc_00E03FE1: mov ecx, var_14
  loc_00E03FE4: pop edi
  loc_00E03FE5: pop esi
  loc_00E03FE6: mov fs:[00000000h], ecx
  loc_00E03FED: pop ebx
  loc_00E03FEE: mov esp, ebp
  loc_00E03FF0: pop ebp
  loc_00E03FF1: retn 0008h
End Sub

Private Sub Proc_2_100_DEB620(arg_C) 'DEB620
  loc_00DEB620: sub esp, 0000000Ch
  loc_00DEB623: mov ecx, arg_1C
  loc_00DEB627: push ebx
  loc_00DEB628: push ebp
  loc_00DEB629: push esi
  loc_00DEB62A: xor eax, eax
  loc_00DEB62C: push edi
  loc_00DEB62D: mov arg_C, eax
  loc_00DEB631: push ecx
  loc_00DEB632: push 00000001h
  loc_00DEB634: push eax
  loc_00DEB635: mov arg_1C, eax
  loc_00DEB639: mov arg_14, eax
  loc_00DEB63D: call 006DD520h ; CreatePen(%x1v, %x2v, %x3v)
  loc_00DEB642: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DEB648: mov Me, eax
  loc_00DEB64C: call edi
  loc_00DEB64E: mov esi, arg_18
  loc_00DEB652: mov ebp, Me
  loc_00DEB656: lea eax, Me
  loc_00DEB65A: mov edx, [esi]
  loc_00DEB65C: push eax
  loc_00DEB65D: push esi
  loc_00DEB65E: call [edx+000000D8h]
  loc_00DEB664: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB66A: test eax, eax
  loc_00DEB66C: fnclex
  loc_00DEB66E: jge 00DEB67Eh
  loc_00DEB670: push 000000D8h
  loc_00DEB675: push 006DCB00h
  loc_00DEB67A: push esi
  loc_00DEB67B: push eax
  loc_00DEB67C: call ebx
  loc_00DEB67E: mov ecx, Me
  loc_00DEB682: push ebp
  loc_00DEB683: push ecx
  loc_00DEB684: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEB689: mov arg_2C, eax
  loc_00DEB68D: call edi
  loc_00DEB68F: mov edx, [esi]
  loc_00DEB691: lea eax, Me
  loc_00DEB695: push eax
  loc_00DEB696: push esi
  loc_00DEB697: call [edx+000000D8h]
  loc_00DEB69D: test eax, eax
  loc_00DEB69F: fnclex
  loc_00DEB6A1: jge 00DEB6B1h
  loc_00DEB6A3: push 000000D8h
  loc_00DEB6A8: push 006DCB00h
  loc_00DEB6AD: push esi
  loc_00DEB6AE: push eax
  loc_00DEB6AF: call ebx
  loc_00DEB6B1: mov edx, arg_20
  loc_00DEB6B5: mov eax, arg_1C
  loc_00DEB6B9: lea ecx, arg_C
  loc_00DEB6BD: push ecx
  loc_00DEB6BE: mov ecx, arg_C
  loc_00DEB6C2: push edx
  loc_00DEB6C3: push eax
  loc_00DEB6C4: push ecx
  loc_00DEB6C5: call 006DD2B0h ; MoveToEx(%x1v, %x2v, %x3v, %x4v)
  loc_00DEB6CA: call edi
  loc_00DEB6CC: mov edx, [esi]
  loc_00DEB6CE: lea eax, Me
  loc_00DEB6D2: push eax
  loc_00DEB6D3: push esi
  loc_00DEB6D4: call [edx+000000D8h]
  loc_00DEB6DA: test eax, eax
  loc_00DEB6DC: fnclex
  loc_00DEB6DE: jge 00DEB6EEh
  loc_00DEB6E0: push 000000D8h
  loc_00DEB6E5: push 006DCB00h
  loc_00DEB6EA: push esi
  loc_00DEB6EB: push eax
  loc_00DEB6EC: call ebx
  loc_00DEB6EE: mov ecx, arg_28
  loc_00DEB6F2: mov edx, arg_24
  loc_00DEB6F6: mov eax, Me
  loc_00DEB6FA: push ecx
  loc_00DEB6FB: push edx
  loc_00DEB6FC: push eax
  loc_00DEB6FD: call 006DD5B8h ; LineTo(%x1v, %x2v, %x3v)
  loc_00DEB702: call edi
  loc_00DEB704: mov ecx, [esi]
  loc_00DEB706: lea edx, Me
  loc_00DEB70A: push edx
  loc_00DEB70B: push esi
  loc_00DEB70C: call [ecx+000000D8h]
  loc_00DEB712: test eax, eax
  loc_00DEB714: fnclex
  loc_00DEB716: jge 00DEB726h
  loc_00DEB718: push 000000D8h
  loc_00DEB71D: push 006DCB00h
  loc_00DEB722: push esi
  loc_00DEB723: push eax
  loc_00DEB724: call ebx
  loc_00DEB726: mov eax, arg_2C
  loc_00DEB72A: mov ecx, Me
  loc_00DEB72E: push eax
  loc_00DEB72F: push ecx
  loc_00DEB730: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEB735: call edi
  loc_00DEB737: push ebp
  loc_00DEB738: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEB73D: call edi
  loc_00DEB73F: pop edi
  loc_00DEB740: pop esi
  loc_00DEB741: pop ebp
  loc_00DEB742: xor eax, eax
  loc_00DEB744: pop ebx
  loc_00DEB745: add esp, 0000000Ch
  loc_00DEB748: retn 0018h
End Sub

Private Sub Proc_2_101_DEB750(arg_10, arg_14, arg_18) 'DEB750
  loc_00DEB750: mov eax, Me
  loc_00DEB754: cmp [eax], 00000000h
  loc_00DEB757: jg 00DEB75Fh
  loc_00DEB759: mov [eax], 00000000h
  loc_00DEB75F: cmp [eax], 00000064h
  loc_00DEB762: jl 00DEB76Ah
  loc_00DEB764: mov [eax], 00000064h
  loc_00DEB76A: mov eax, [esp+00000008h]
  loc_00DEB76E: push ebx
  loc_00DEB76F: push ebp
  loc_00DEB770: push esi
  loc_00DEB771: push edi
  loc_00DEB772: mov edi, [eax]
  loc_00DEB774: mov eax, edi
  loc_00DEB776: mov ecx, arg_14
  loc_00DEB77A: cdq
  loc_00DEB77B: and edx, 000000FFh
  loc_00DEB781: mov ecx, [ecx]
  loc_00DEB783: add eax, edx
  loc_00DEB785: mov ebx, edi
  loc_00DEB787: mov esi, eax
  loc_00DEB789: mov eax, edi
  loc_00DEB78B: cdq
  loc_00DEB78C: and edx, 0000FFFFh
  loc_00DEB792: and ebx, 000000FFh
  loc_00DEB798: add eax, edx
  loc_00DEB79A: mov edi, eax
  loc_00DEB79C: mov eax, ecx
  loc_00DEB79E: and eax, 000000FFh
  loc_00DEB7A3: mov arg_10, eax
  loc_00DEB7A7: mov eax, ecx
  loc_00DEB7A9: cdq
  loc_00DEB7AA: and edx, 000000FFh
  loc_00DEB7B0: add eax, edx
  loc_00DEB7B2: mov ebp, eax
  loc_00DEB7B4: mov eax, ecx
  loc_00DEB7B6: cdq
  loc_00DEB7B7: and edx, 0000FFFFh
  loc_00DEB7BD: mov ecx, ebx
  loc_00DEB7BF: add eax, edx
  loc_00DEB7C1: mov edx, arg_18
  loc_00DEB7C5: sar eax, 10h
  loc_00DEB7C8: and eax, 000000FFh
  loc_00DEB7CD: mov arg_14, eax
  loc_00DEB7D1: mov eax, [edx]
  loc_00DEB7D3: mov edx, arg_10
  loc_00DEB7D7: mov arg_18, eax
  loc_00DEB7DB: sar esi, 08h
  loc_00DEB7DE: sar edi, 10h
  loc_00DEB7E1: sar ebp, 08h
  loc_00DEB7E4: and esi, 000000FFh
  loc_00DEB7EA: and edi, 000000FFh
  loc_00DEB7F0: and ebp, 000000FFh
  loc_00DEB7F6: sub ecx, edx
  loc_00DEB7F8: jo 00DEB897h
  loc_00DEB7FE: imul ecx, eax
  loc_00DEB801: mov eax, 51EB851Fh
  loc_00DEB806: jo 00DEB897h
  loc_00DEB80C: imul ecx
  loc_00DEB80E: sar edx, 05h
  loc_00DEB811: mov eax, edx
  loc_00DEB813: mov ecx, esi
  loc_00DEB815: shr eax, 1Fh
  loc_00DEB818: add edx, eax
  loc_00DEB81A: mov eax, 51EB851Fh
  loc_00DEB81F: add edx, ebx
  loc_00DEB821: mov ebx, arg_18
  loc_00DEB825: jo 00DEB897h
  loc_00DEB827: sub ecx, ebp
  loc_00DEB829: mov arg_10, edx
  loc_00DEB82D: jo 00DEB897h
  loc_00DEB82F: imul ecx, ebx
  loc_00DEB832: jo 00DEB897h
  loc_00DEB834: imul ecx
  loc_00DEB836: mov ecx, edx
  loc_00DEB838: mov ebp, arg_14
  loc_00DEB83C: sar ecx, 05h
  loc_00DEB83F: mov edx, ecx
  loc_00DEB841: mov eax, 51EB851Fh
  loc_00DEB846: shr edx, 1Fh
  loc_00DEB849: add ecx, edx
  loc_00DEB84B: mov edx, edi
  loc_00DEB84D: add ecx, esi
  loc_00DEB84F: jo 00DEB897h
  loc_00DEB851: sub edx, ebp
  loc_00DEB853: jo 00DEB897h
  loc_00DEB855: imul edx, ebx
  loc_00DEB858: jo 00DEB897h
  loc_00DEB85A: imul edx
  loc_00DEB85C: mov eax, edx
  loc_00DEB85E: mov esi, arg_10
  loc_00DEB862: sar eax, 05h
  loc_00DEB865: mov edx, eax
  loc_00DEB867: shr edx, 1Fh
  loc_00DEB86A: add eax, edx
  loc_00DEB86C: add eax, edi
  loc_00DEB86E: pop edi
  loc_00DEB86F: jo 00DEB897h
  loc_00DEB871: imul ecx, ecx, 00000100h
  loc_00DEB877: jo 00DEB897h
  loc_00DEB879: add ecx, esi
  loc_00DEB87B: pop esi
  loc_00DEB87C: jo 00DEB897h
  loc_00DEB87E: imul eax, eax, 00010000h
  loc_00DEB884: jo 00DEB897h
  loc_00DEB886: add ecx, eax
  loc_00DEB888: mov eax, arg_14
  loc_00DEB88C: jo 00DEB897h
  loc_00DEB88E: mov [eax], ecx
  loc_00DEB890: pop ebp
  loc_00DEB891: xor eax, eax
  loc_00DEB893: pop ebx
  loc_00DEB894: retn 0014h
End Sub

Private Sub Proc_2_102_DEB8A0(arg_C, arg_10, arg_14, arg_18, arg_1C) 'DEB8A0
  loc_00DEB8A0: sub esp, 00000008h
  loc_00DEB8A3: push ebx
  loc_00DEB8A4: push esi
  loc_00DEB8A5: mov esi, arg_10
  loc_00DEB8A9: push edi
  loc_00DEB8AA: mov eax, esi
  loc_00DEB8AC: mov edi, arg_18
  loc_00DEB8B0: cdq
  loc_00DEB8B1: and edx, 0000FFFFh
  loc_00DEB8B7: mov ebx, [00401200h] ; __vbaFpI2
  loc_00DEB8BD: add eax, edx
  loc_00DEB8BF: mov ecx, eax
  loc_00DEB8C1: mov eax, edi
  loc_00DEB8C3: cdq
  loc_00DEB8C4: and edx, 0000FFFFh
  loc_00DEB8CA: add eax, edx
  loc_00DEB8CC: sar ecx, 10h
  loc_00DEB8CF: sar eax, 10h
  loc_00DEB8D2: and ecx, 000000FFh
  loc_00DEB8D8: and eax, 000000FFh
  loc_00DEB8DD: add ecx, eax
  loc_00DEB8DF: jo 00DEB9F7h
  loc_00DEB8E5: mov arg_14, ecx
  loc_00DEB8E9: fild real4 ptr arg_14
  loc_00DEB8ED: fstp real8 ptr Proc_2_102_DEB8A0
  loc_00DEB8F1: fld real8 ptr Proc_2_102_DEB8A0
  loc_00DEB8F5: cmp [00E3D000h], 00000000h
  loc_00DEB8FC: jnz 00DEB906h
  loc_00DEB8FE: fdiv st0, real8 ptr [00401338h]
  loc_00DEB904: jmp 00DEB917h
  loc_00DEB906: push [0040133Ch]
  loc_00DEB90C: push [00401338h]
  loc_00DEB912: call 00402854h ; _adj_fdiv_m64
  loc_00DEB917: fnstsw ax
  loc_00DEB919: test al, 0Dh
  loc_00DEB91B: jnz 00DEB9F2h
  loc_00DEB921: call ebx
  loc_00DEB923: push eax
  loc_00DEB924: mov eax, esi
  loc_00DEB926: cdq
  loc_00DEB927: and edx, 000000FFh
  loc_00DEB92D: add eax, edx
  loc_00DEB92F: mov ecx, eax
  loc_00DEB931: mov eax, edi
  loc_00DEB933: cdq
  loc_00DEB934: and edx, 000000FFh
  loc_00DEB93A: add eax, edx
  loc_00DEB93C: sar ecx, 08h
  loc_00DEB93F: sar eax, 08h
  loc_00DEB942: and ecx, 000000FFh
  loc_00DEB948: and eax, 000000FFh
  loc_00DEB94D: add ecx, eax
  loc_00DEB94F: jo 00DEB9F7h
  loc_00DEB955: mov arg_18, ecx
  loc_00DEB959: fild real4 ptr arg_18
  loc_00DEB95D: fstp real8 ptr Me
  loc_00DEB961: fld real8 ptr Me
  loc_00DEB965: cmp [00E3D000h], 00000000h
  loc_00DEB96C: jnz 00DEB976h
  loc_00DEB96E: fdiv st0, real8 ptr [00401338h]
  loc_00DEB974: jmp 00DEB987h
  loc_00DEB976: push [0040133Ch]
  loc_00DEB97C: push [00401338h]
  loc_00DEB982: call 00402854h ; _adj_fdiv_m64
  loc_00DEB987: fnstsw ax
  loc_00DEB989: test al, 0Dh
  loc_00DEB98B: jnz 00DEB9F2h
  loc_00DEB98D: call ebx
  loc_00DEB98F: and esi, 000000FFh
  loc_00DEB995: and edi, 000000FFh
  loc_00DEB99B: add esi, edi
  loc_00DEB99D: push eax
  loc_00DEB99E: jo 00DEB9F7h
  loc_00DEB9A0: mov arg_1C, esi
  loc_00DEB9A4: fild real4 ptr arg_1C
  loc_00DEB9A8: fstp real8 ptr arg_C
  loc_00DEB9AC: fld real8 ptr arg_C
  loc_00DEB9B0: cmp [00E3D000h], 00000000h
  loc_00DEB9B7: jnz 00DEB9C1h
  loc_00DEB9B9: fdiv st0, real8 ptr [00401338h]
  loc_00DEB9BF: jmp 00DEB9D2h
  loc_00DEB9C1: push [0040133Ch]
  loc_00DEB9C7: push [00401338h]
  loc_00DEB9CD: call 00402854h ; _adj_fdiv_m64
  loc_00DEB9D2: fnstsw ax
  loc_00DEB9D4: test al, 0Dh
  loc_00DEB9D6: jnz 00DEB9F2h
  loc_00DEB9D8: call ebx
  loc_00DEB9DA: push eax
  loc_00DEB9DB: call [00401028h] ; rtcRgb
  loc_00DEB9E1: pop edi
  loc_00DEB9E2: pop esi
  loc_00DEB9E3: mov edx, arg_14
  loc_00DEB9E7: pop ebx
  loc_00DEB9E8: mov [edx], eax
  loc_00DEB9EA: xor eax, eax
  loc_00DEB9EC: add esp, 00000008h
  loc_00DEB9EF: retn 0010h
End Sub

Private Sub Proc_2_103_DEBA00(arg_C, arg_10, arg_14, arg_18, arg_1C) 'DEBA00
  loc_00DEBA00: sub esp, 00000014h
  loc_00DEBA03: xor eax, eax
  loc_00DEBA05: mov ecx, arg_18
  loc_00DEBA09: mov var_4, eax
  loc_00DEBA0D: push ebx
  loc_00DEBA0E: mov Proc_2_103_DEBA00, eax
  loc_00DEBA12: push esi
  loc_00DEBA13: mov esi, arg_28
  loc_00DEBA17: mov arg_C, eax
  loc_00DEBA1B: mov arg_10, eax
  loc_00DEBA1F: mov [esp+00000008h], eax
  loc_00DEBA23: mov eax, arg_1C
  loc_00DEBA27: push edi
  loc_00DEBA28: mov edi, arg_28
  loc_00DEBA2C: mov Me, eax
  loc_00DEBA30: add eax, edi
  loc_00DEBA32: mov arg_C, ecx
  loc_00DEBA36: jo 00DEBAB1h
  loc_00DEBA38: add ecx, esi
  loc_00DEBA3A: mov arg_10, eax
  loc_00DEBA3E: jo 00DEBAB1h
  loc_00DEBA40: mov arg_14, ecx
  loc_00DEBA44: mov ecx, arg_30
  loc_00DEBA48: push ecx
  loc_00DEBA49: call 006DD26Ch ; CreateSolidBrush(%x1v)
  loc_00DEBA4E: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DEBA54: mov Proc_2_103_DEBA00, eax
  loc_00DEBA58: call edi
  loc_00DEBA5A: mov esi, arg_1C
  loc_00DEBA5E: mov ebx, Proc_2_103_DEBA00
  loc_00DEBA62: lea eax, Proc_2_103_DEBA00
  loc_00DEBA66: mov edx, [esi]
  loc_00DEBA68: push eax
  loc_00DEBA69: push esi
  loc_00DEBA6A: call [edx+000000D8h]
  loc_00DEBA70: test eax, eax
  loc_00DEBA72: fnclex
  loc_00DEBA74: jge 00DEBA88h
  loc_00DEBA76: push 000000D8h
  loc_00DEBA7B: push 006DCB00h
  loc_00DEBA80: push esi
  loc_00DEBA81: push eax
  loc_00DEBA82: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEBA88: mov edx, Proc_2_103_DEBA00
  loc_00DEBA8C: lea ecx, Me
  loc_00DEBA90: push ebx
  loc_00DEBA91: push ecx
  loc_00DEBA92: push edx
  loc_00DEBA93: call 006DDB10h ; FrameRect(%x1v, %x2v, %x3v)
  loc_00DEBA98: call edi
  loc_00DEBA9A: push ebx
  loc_00DEBA9B: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEBAA0: mov Proc_2_103_DEBA00, eax
  loc_00DEBAA4: call edi
  loc_00DEBAA6: pop edi
  loc_00DEBAA7: pop esi
  loc_00DEBAA8: xor eax, eax
  loc_00DEBAAA: pop ebx
  loc_00DEBAAB: add esp, 00000014h
  loc_00DEBAAE: retn 0018h
End Sub

Private Sub Proc_2_104_DEBAC0(arg_C, arg_10, arg_14, arg_18) 'DEBAC0
  loc_00DEBAC0: sub esp, 00000014h
  loc_00DEBAC3: xor eax, eax
  loc_00DEBAC5: mov ecx, arg_18
  loc_00DEBAC9: mov var_4, eax
  loc_00DEBACD: mov edx, arg_20
  loc_00DEBAD1: mov [esp+00000008h], eax
  loc_00DEBAD5: push esi
  loc_00DEBAD6: mov esi, arg_20
  loc_00DEBADA: mov Me, eax
  loc_00DEBADE: mov arg_C, eax
  loc_00DEBAE2: mov var_4, eax
  loc_00DEBAE6: mov eax, arg_18
  loc_00DEBAEA: mov Proc_2_104_DEBAC0, ecx
  loc_00DEBAEE: mov [esp+00000008h], eax
  loc_00DEBAF2: add eax, esi
  loc_00DEBAF4: mov esi, arg_14
  loc_00DEBAF8: jo 00DEBB4Ah
  loc_00DEBAFA: add ecx, edx
  loc_00DEBAFC: lea edx, var_4
  loc_00DEBB00: jo 00DEBB4Ah
  loc_00DEBB02: mov arg_C, ecx
  loc_00DEBB06: mov ecx, [esi]
  loc_00DEBB08: push edx
  loc_00DEBB09: push esi
  loc_00DEBB0A: mov arg_10, eax
  loc_00DEBB0E: call [ecx+000000D8h]
  loc_00DEBB14: test eax, eax
  loc_00DEBB16: fnclex
  loc_00DEBB18: jge 00DEBB2Ch
  loc_00DEBB1A: push 000000D8h
  loc_00DEBB1F: push 006DCB00h
  loc_00DEBB24: push esi
  loc_00DEBB25: push eax
  loc_00DEBB26: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEBB2C: mov ecx, var_4
  loc_00DEBB30: lea eax, [esp+00000008h]
  loc_00DEBB34: push eax
  loc_00DEBB35: push ecx
  loc_00DEBB36: call 006DDA18h ; DrawFocusRect(%x1v, %x2v)
  loc_00DEBB3B: call [00401074h] ; __vbaSetSystemError
  loc_00DEBB41: xor eax, eax
  loc_00DEBB43: pop esi
  loc_00DEBB44: add esp, 00000014h
  loc_00DEBB47: retn 0014h
End Sub

Private Function Proc_2_105_DEBB50(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28, arg_2C, arg_30) 'DEBB50
  loc_00DEBB50: push ebp
  loc_00DEBB51: mov ebp, esp
  loc_00DEBB53: sub esp, 00000008h
  loc_00DEBB56: push 00402836h ; __vbaExceptHandler
  loc_00DEBB5B: mov eax, fs:[00000000h]
  loc_00DEBB61: push eax
  loc_00DEBB62: mov fs:[00000000h], esp
  loc_00DEBB69: sub esp, 00000184h
  loc_00DEBB6F: push ebx
  loc_00DEBB70: push esi
  loc_00DEBB71: push edi
  loc_00DEBB72: mov var_8, esp
  loc_00DEBB75: mov var_4, 00401360h
  loc_00DEBB7C: mov ecx, 0000000Bh
  loc_00DEBB81: xor eax, eax
  loc_00DEBB83: lea edi, var_6C
  loc_00DEBB86: lea edx, var_24
  loc_00DEBB89: repz stosd
  loc_00DEBB8B: mov ecx, arg_20
  loc_00DEBB8E: xor esi, esi
  loc_00DEBB90: mov var_78, ax
  loc_00DEBB94: push ecx
  loc_00DEBB95: push edx
  loc_00DEBB96: mov var_18, 00h
  loc_00DEBB9A: mov var_24, esi
  loc_00DEBB9D: mov var_3C, esi
  loc_00DEBBA0: mov var_40, esi
  loc_00DEBBA3: mov var_76, al
  loc_00DEBBA6: mov var_90, al
  loc_00DEBBAC: mov var_A0, esi
  loc_00DEBBB2: mov var_A8, esi
  loc_00DEBBB8: mov var_B8, esi
  loc_00DEBBBE: mov var_C8, esi
  loc_00DEBBC4: mov var_D8, esi
  loc_00DEBBCA: mov var_E8, esi
  loc_00DEBBD0: mov var_EC, esi
  loc_00DEBBD6: mov var_F0, esi
  loc_00DEBBDC: mov var_108, esi
  loc_00DEBBE2: call [004010B4h] ; __vbaObjSetAddref
  loc_00DEBBE8: cmp arg_18, esi
  loc_00DEBBEB: jz 00DED477h
  loc_00DEBBF1: mov edi, arg_1C
  loc_00DEBBF4: cmp edi, esi
  loc_00DEBBF6: jz 00DED477h
  loc_00DEBBFC: mov eax, var_24
  loc_00DEBBFF: push esi
  loc_00DEBC00: push eax
  loc_00DEBC01: call [0040114Ch] ; __vbaObjIs
  loc_00DEBC07: test ax, ax
  loc_00DEBC0A: jnz 00DED477h
  loc_00DEBC10: mov ebx, Me
  loc_00DEBC13: mov eax, [ebx+00000048h]
  loc_00DEBC16: cmp eax, 00000001h
  loc_00DEBC19: jnz 00DEBC26h
  loc_00DEBC1B: mov ecx, [ebx+00000168h]
  loc_00DEBC21: mov var_3C, ecx
  loc_00DEBC24: jmp 00DEBC34h
  loc_00DEBC26: cmp eax, 00000002h
  loc_00DEBC29: jnz 00DEBC34h
  loc_00DEBC2B: mov edx, [ebx+0000016Ch]
  loc_00DEBC31: mov var_3C, edx
  loc_00DEBC34: cmp [ebx+0000007Ch], si
  loc_00DEBC38: jnz 00DEBC9Dh
  loc_00DEBC3A: mov eax, [ebx+0000015Ch]
  loc_00DEBC40: sub eax, esi
  loc_00DEBC42: jz 00DEBC8Ch
  loc_00DEBC44: dec eax
  loc_00DEBC45: jnz 00DEBC9Dh
  loc_00DEBC47: xor eax, eax
  loc_00DEBC49: mov al, [ebx+00000158h]
  loc_00DEBC4F: mov var_12C, eax
  loc_00DEBC55: fild real4 ptr var_12C
  loc_00DEBC5B: fstp real8 ptr var_134
  loc_00DEBC61: fld real8 ptr var_134
  loc_00DEBC67: fmul st0, real8 ptr [00401358h]
  loc_00DEBC6D: fnstsw ax
  loc_00DEBC6F: test al, 0Dh
  loc_00DEBC71: jnz 00DED4DEh
  loc_00DEBC77: call [0040111Ch] ; __vbaFpUI1
  loc_00DEBC7D: mov var_90, al
  loc_00DEBC83: mov arg_30, FFFFFFFFh
  loc_00DEBC8A: jmp 00DEBC9Dh
  loc_00DEBC8C: mov ecx, 00000034h
  loc_00DEBC91: call [00401144h] ; __vbaUI1I2
  loc_00DEBC97: mov var_90, al
  loc_00DEBC9D: cmp [ebx+00000048h], 00000001h
  loc_00DEBCA1: jnz 00DEBCACh
  loc_00DEBCA3: mov cl, [ebx+00000159h]
  loc_00DEBCA9: mov var_18, cl
  loc_00DEBCAC: mov edx, [ebx]
  loc_00DEBCAE: lea eax, var_F0
  loc_00DEBCB4: push eax
  loc_00DEBCB5: push ebx
  loc_00DEBCB6: call [edx+000000D8h]
  loc_00DEBCBC: cmp eax, esi
  loc_00DEBCBE: fnclex
  loc_00DEBCC0: jge 00DEBCD4h
  loc_00DEBCC2: push 000000D8h
  loc_00DEBCC7: push 006DCB00h
  loc_00DEBCCC: push ebx
  loc_00DEBCCD: push eax
  loc_00DEBCCE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEBCD4: mov ecx, var_F0
  loc_00DEBCDA: push ecx
  loc_00DEBCDB: call 006DD3F0h ; CreateCompatibleDC(%x1v)
  loc_00DEBCE0: mov esi, [00401074h] ; __vbaSetSystemError
  loc_00DEBCE6: mov var_F4, eax
  loc_00DEBCEC: call __vbaSetSystemError
  loc_00DEBCEE: mov eax, arg_18
  loc_00DEBCF1: mov edx, var_F4
  loc_00DEBCF7: test eax, eax
  loc_00DEBCF9: mov var_94, edx
  loc_00DEBCFF: jge 00DEBE30h
  loc_00DEBD05: mov eax, var_24
  loc_00DEBD08: push 00000000h
  loc_00DEBD0A: push 00000004h
  loc_00DEBD0C: lea ecx, var_B8
  loc_00DEBD12: push eax
  loc_00DEBD13: push ecx
  loc_00DEBD14: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DEBD1A: mov eax, [ebx+00000010h]
  loc_00DEBD1D: add esp, 00000010h
  loc_00DEBD20: lea ecx, var_EC
  loc_00DEBD26: mov edx, [eax]
  loc_00DEBD28: push ecx
  loc_00DEBD29: push eax
  loc_00DEBD2A: call [edx+00000110h]
  loc_00DEBD30: test eax, eax
  loc_00DEBD32: fnclex
  loc_00DEBD34: jge 00DEBD4Bh
  loc_00DEBD36: mov edx, [ebx+00000010h]
  loc_00DEBD39: push 00000110h
  loc_00DEBD3E: push 006DCB00h
  loc_00DEBD43: push edx
  loc_00DEBD44: push eax
  loc_00DEBD45: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEBD4B: mov ax, var_EC
  loc_00DEBD52: mov edx, [ebx+00000010h]
  loc_00DEBD55: mov var_D0, ax
  loc_00DEBD5C: mov eax, 00000002h
  loc_00DEBD61: mov ecx, 00000008h
  loc_00DEBD66: mov var_D8, eax
  loc_00DEBD6C: mov var_C0, ecx
  loc_00DEBD72: mov var_C8, eax
  loc_00DEBD78: mov ebx, [edx]
  loc_00DEBD7A: lea edx, var_F0
  loc_00DEBD80: push edx
  loc_00DEBD81: sub esp, 00000010h
  loc_00DEBD84: mov edx, esp
  loc_00DEBD86: sub esp, 00000010h
  loc_00DEBD89: mov [edx], eax
  loc_00DEBD8B: mov eax, var_D4
  loc_00DEBD91: mov [edx+00000004h], eax
  loc_00DEBD94: mov eax, var_D0
  loc_00DEBD9A: mov [edx+00000008h], eax
  loc_00DEBD9D: mov eax, var_CC
  loc_00DEBDA3: mov [edx+0000000Ch], eax
  loc_00DEBDA6: mov eax, var_C8
  loc_00DEBDAC: mov edx, esp
  loc_00DEBDAE: mov [edx], eax
  loc_00DEBDB0: mov eax, var_C4
  loc_00DEBDB6: mov [edx+00000004h], eax
  loc_00DEBDB9: mov [edx+00000008h], ecx
  loc_00DEBDBC: mov ecx, var_BC
  loc_00DEBDC2: mov [edx+0000000Ch], ecx
  loc_00DEBDC5: lea edx, var_B8
  loc_00DEBDCB: push edx
  loc_00DEBDCC: call [004011D0h] ; __vbaI4Var
  loc_00DEBDD2: mov var_138, eax
  loc_00DEBDD8: mov edx, ebx
  loc_00DEBDDA: fild real4 ptr var_138
  loc_00DEBDE0: mov ebx, Me
  loc_00DEBDE3: fstp real4 ptr var_13C
  loc_00DEBDE9: mov eax, var_13C
  loc_00DEBDEF: mov ecx, [ebx+00000010h]
  loc_00DEBDF2: push eax
  loc_00DEBDF3: push ecx
  loc_00DEBDF4: call [edx+00000374h]
  loc_00DEBDFA: test eax, eax
  loc_00DEBDFC: fnclex
  loc_00DEBDFE: jge 00DEBE15h
  loc_00DEBE00: mov ecx, [ebx+00000010h]
  loc_00DEBE03: push 00000374h
  loc_00DEBE08: push 006DCB00h
  loc_00DEBE0D: push ecx
  loc_00DEBE0E: push eax
  loc_00DEBE0F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEBE15: fld real4 ptr var_F0
  loc_00DEBE1B: call [00401208h] ; __vbaFpI4
  loc_00DEBE21: lea ecx, var_B8
  loc_00DEBE27: mov arg_18, eax
  loc_00DEBE2A: call [00401024h] ; __vbaFreeVar
  loc_00DEBE30: test edi, edi
  loc_00DEBE32: jge 00DEBF60h
  loc_00DEBE38: mov edx, var_24
  loc_00DEBE3B: push 00000000h
  loc_00DEBE3D: push 00000005h
  loc_00DEBE3F: lea eax, var_B8
  loc_00DEBE45: push edx
  loc_00DEBE46: push eax
  loc_00DEBE47: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DEBE4D: mov eax, [ebx+00000010h]
  loc_00DEBE50: add esp, 00000010h
  loc_00DEBE53: lea edx, var_EC
  loc_00DEBE59: mov ecx, [eax]
  loc_00DEBE5B: push edx
  loc_00DEBE5C: push eax
  loc_00DEBE5D: call [ecx+00000110h]
  loc_00DEBE63: test eax, eax
  loc_00DEBE65: fnclex
  loc_00DEBE67: jge 00DEBE7Eh
  loc_00DEBE69: mov ecx, [ebx+00000010h]
  loc_00DEBE6C: push 00000110h
  loc_00DEBE71: push 006DCB00h
  loc_00DEBE76: push ecx
  loc_00DEBE77: push eax
  loc_00DEBE78: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEBE7E: mov dx, var_EC
  loc_00DEBE85: mov eax, 00000002h
  loc_00DEBE8A: mov var_D0, dx
  loc_00DEBE91: mov edx, [ebx+00000010h]
  loc_00DEBE94: mov ecx, 00000008h
  loc_00DEBE99: mov var_D8, eax
  loc_00DEBE9F: mov var_C0, ecx
  loc_00DEBEA5: mov var_C8, eax
  loc_00DEBEAB: mov edi, [edx]
  loc_00DEBEAD: lea edx, var_F0
  loc_00DEBEB3: push edx
  loc_00DEBEB4: sub esp, 00000010h
  loc_00DEBEB7: mov edx, esp
  loc_00DEBEB9: sub esp, 00000010h
  loc_00DEBEBC: mov [edx], eax
  loc_00DEBEBE: mov eax, var_D4
  loc_00DEBEC4: mov [edx+00000004h], eax
  loc_00DEBEC7: mov eax, var_D0
  loc_00DEBECD: mov [edx+00000008h], eax
  loc_00DEBED0: mov eax, var_CC
  loc_00DEBED6: mov [edx+0000000Ch], eax
  loc_00DEBED9: mov eax, var_C8
  loc_00DEBEDF: mov edx, esp
  loc_00DEBEE1: mov [edx], eax
  loc_00DEBEE3: mov eax, var_C4
  loc_00DEBEE9: mov [edx+00000004h], eax
  loc_00DEBEEC: mov [edx+00000008h], ecx
  loc_00DEBEEF: mov ecx, var_BC
  loc_00DEBEF5: mov [edx+0000000Ch], ecx
  loc_00DEBEF8: lea edx, var_B8
  loc_00DEBEFE: push edx
  loc_00DEBEFF: call [004011D0h] ; __vbaI4Var
  loc_00DEBF05: mov var_144, eax
  loc_00DEBF0B: mov ecx, [ebx+00000010h]
  loc_00DEBF0E: fild real4 ptr var_144
  loc_00DEBF14: fstp real4 ptr var_148
  loc_00DEBF1A: mov eax, var_148
  loc_00DEBF20: push eax
  loc_00DEBF21: push ecx
  loc_00DEBF22: call [edi+00000378h]
  loc_00DEBF28: test eax, eax
  loc_00DEBF2A: fnclex
  loc_00DEBF2C: jge 00DEBF43h
  loc_00DEBF2E: mov edx, [ebx+00000010h]
  loc_00DEBF31: push 00000378h
  loc_00DEBF36: push 006DCB00h
  loc_00DEBF3B: push edx
  loc_00DEBF3C: push eax
  loc_00DEBF3D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEBF43: fld real4 ptr var_F0
  loc_00DEBF49: call [00401208h] ; __vbaFpI4
  loc_00DEBF4F: mov edi, eax
  loc_00DEBF51: lea ecx, var_B8
  loc_00DEBF57: mov arg_1C, edi
  loc_00DEBF5A: call [00401024h] ; __vbaFreeVar
  loc_00DEBF60: mov eax, var_24
  loc_00DEBF63: push 00000000h
  loc_00DEBF65: push 00000003h
  loc_00DEBF67: lea ecx, var_B8
  loc_00DEBF6D: push eax
  loc_00DEBF6E: push ecx
  loc_00DEBF6F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DEBF75: add esp, 00000010h
  loc_00DEBF78: push eax
  loc_00DEBF79: call [0040118Ch] ; __vbaI2Var
  loc_00DEBF7F: dec ax
  loc_00DEBF81: lea ecx, var_B8
  loc_00DEBF87: neg ax
  loc_00DEBF8A: sbb eax, eax
  loc_00DEBF8C: inc eax
  loc_00DEBF8D: neg eax
  loc_00DEBF8F: mov var_F8, eax
  loc_00DEBF95: call [00401024h] ; __vbaFreeVar
  loc_00DEBF9B: cmp var_F8, 0000h
  loc_00DEBFA3: jz 00DEBFF6h
  loc_00DEBFA5: mov edx, var_24
  loc_00DEBFA8: push 00000000h
  loc_00DEBFAA: push 00000000h
  loc_00DEBFAC: lea eax, var_B8
  loc_00DEBFB2: push edx
  loc_00DEBFB3: push eax
  loc_00DEBFB4: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DEBFBA: add esp, 00000010h
  loc_00DEBFBD: push eax
  loc_00DEBFBE: call [004011D0h] ; __vbaI4Var
  loc_00DEBFC4: mov ecx, var_F4
  loc_00DEBFCA: push eax
  loc_00DEBFCB: push ecx
  loc_00DEBFCC: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEBFD1: mov var_F0, eax
  loc_00DEBFD7: call __vbaSetSystemError
  loc_00DEBFD9: mov edx, var_F0
  loc_00DEBFDF: lea ecx, var_B8
  loc_00DEBFE5: mov var_9C, edx
  loc_00DEBFEB: call [00401024h] ; __vbaFreeVar
  loc_00DEBFF1: jmp 00DEC0A9h
  loc_00DEBFF6: mov eax, arg_18
  loc_00DEBFF9: mov ecx, arg_C
  loc_00DEBFFC: push edi
  loc_00DEBFFD: push eax
  loc_00DEBFFE: push ecx
  loc_00DEBFFF: call 006DD440h ; CreateCompatibleBitmap(%x1v, %x2v, %x3v)
  loc_00DEC004: mov var_F0, eax
  loc_00DEC00A: call __vbaSetSystemError
  loc_00DEC00C: mov edx, var_F0
  loc_00DEC012: mov eax, var_F4
  loc_00DEC018: push edx
  loc_00DEC019: push eax
  loc_00DEC01A: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEC01F: mov var_F4, eax
  loc_00DEC025: call __vbaSetSystemError
  loc_00DEC027: mov edx, arg_24
  loc_00DEC02A: mov ecx, var_F4
  loc_00DEC030: push edx
  loc_00DEC031: mov var_9C, ecx
  loc_00DEC037: call 006DD26Ch ; CreateSolidBrush(%x1v)
  loc_00DEC03C: mov var_F0, eax
  loc_00DEC042: call __vbaSetSystemError
  loc_00DEC044: mov ecx, var_24
  loc_00DEC047: mov eax, var_F0
  loc_00DEC04D: push 00000000h
  loc_00DEC04F: push 00000000h
  loc_00DEC051: lea edx, var_B8
  loc_00DEC057: push ecx
  loc_00DEC058: push edx
  loc_00DEC059: mov var_70, eax
  loc_00DEC05C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DEC062: mov eax, var_70
  loc_00DEC065: mov ecx, arg_18
  loc_00DEC068: add esp, 00000010h
  loc_00DEC06B: lea edx, var_B8
  loc_00DEC071: push 00000003h
  loc_00DEC073: push eax
  loc_00DEC074: push 00000000h
  loc_00DEC076: push edi
  loc_00DEC077: push ecx
  loc_00DEC078: push edx
  loc_00DEC079: call [004011D0h] ; __vbaI4Var
  loc_00DEC07F: push eax
  loc_00DEC080: mov eax, var_94
  loc_00DEC086: push 00000000h
  loc_00DEC088: push 00000000h
  loc_00DEC08A: push eax
  loc_00DEC08B: call 006DE344h ; DrawIconEx(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v)
  loc_00DEC090: call __vbaSetSystemError
  loc_00DEC092: lea ecx, var_B8
  loc_00DEC098: call [00401024h] ; __vbaFreeVar
  loc_00DEC09E: mov ecx, var_70
  loc_00DEC0A1: push ecx
  loc_00DEC0A2: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEC0A7: call __vbaSetSystemError
  loc_00DEC0A9: mov edx, var_94
  loc_00DEC0AF: push edx
  loc_00DEC0B0: call 006DD3F0h ; CreateCompatibleDC(%x1v)
  loc_00DEC0B5: mov var_F0, eax
  loc_00DEC0BB: call __vbaSetSystemError
  loc_00DEC0BD: mov ecx, var_94
  loc_00DEC0C3: mov eax, var_F0
  loc_00DEC0C9: push ecx
  loc_00DEC0CA: mov var_88, eax
  loc_00DEC0D0: call 006DD3F0h ; CreateCompatibleDC(%x1v)
  loc_00DEC0D5: mov var_F0, eax
  loc_00DEC0DB: call __vbaSetSystemError
  loc_00DEC0DD: mov eax, arg_18
  loc_00DEC0E0: mov ecx, arg_C
  loc_00DEC0E3: mov edx, var_F0
  loc_00DEC0E9: push edi
  loc_00DEC0EA: push eax
  loc_00DEC0EB: push ecx
  loc_00DEC0EC: mov var_38, edx
  loc_00DEC0EF: call 006DD440h ; CreateCompatibleBitmap(%x1v, %x2v, %x3v)
  loc_00DEC0F4: mov var_F0, eax
  loc_00DEC0FA: call __vbaSetSystemError
  loc_00DEC0FC: mov eax, arg_18
  loc_00DEC0FF: mov ecx, arg_C
  loc_00DEC102: mov edx, var_F0
  loc_00DEC108: push edi
  loc_00DEC109: push eax
  loc_00DEC10A: push ecx
  loc_00DEC10B: mov var_7C, edx
  loc_00DEC10E: call 006DD440h ; CreateCompatibleBitmap(%x1v, %x2v, %x3v)
  loc_00DEC113: mov var_F0, eax
  loc_00DEC119: call __vbaSetSystemError
  loc_00DEC11B: mov eax, var_7C
  loc_00DEC11E: mov ecx, var_88
  loc_00DEC124: mov edx, var_F0
  loc_00DEC12A: push eax
  loc_00DEC12B: push ecx
  loc_00DEC12C: mov var_84, edx
  loc_00DEC132: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEC137: mov var_F0, eax
  loc_00DEC13D: call __vbaSetSystemError
  loc_00DEC13F: mov eax, var_84
  loc_00DEC145: mov ecx, var_38
  loc_00DEC148: mov edx, var_F0
  loc_00DEC14E: push eax
  loc_00DEC14F: push ecx
  loc_00DEC150: mov var_20, edx
  loc_00DEC153: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEC158: mov var_F0, eax
  loc_00DEC15E: call __vbaSetSystemError
  loc_00DEC160: mov eax, arg_18
  loc_00DEC163: mov edx, var_F0
  loc_00DEC169: imul eax, edi
  loc_00DEC16C: jo 00DED4E3h
  loc_00DEC172: imul eax, eax, 00000003h
  loc_00DEC175: jo 00DED4E3h
  loc_00DEC17B: sub eax, 00000001h
  loc_00DEC17E: push 00000000h
  loc_00DEC180: jo 00DED4E3h
  loc_00DEC186: push eax
  loc_00DEC187: push 00000001h
  loc_00DEC189: lea ecx, var_40
  loc_00DEC18C: push 00000000h
  loc_00DEC18E: push ecx
  loc_00DEC18F: push 00000003h
  loc_00DEC191: push 00000000h
  loc_00DEC193: mov var_30, edx
  loc_00DEC196: call [00401138h] ; __vbaRedim
  loc_00DEC19C: mov edx, var_40
  loc_00DEC19F: add esp, 0000001Ch
  loc_00DEC1A2: push 00000000h
  loc_00DEC1A4: push edx
  loc_00DEC1A5: push 00000001h
  loc_00DEC1A7: call [0040117Ch] ; __vbaUbound
  loc_00DEC1AD: push eax
  loc_00DEC1AE: push 00000001h
  loc_00DEC1B0: lea eax, var_A0
  loc_00DEC1B6: push 00000000h
  loc_00DEC1B8: push eax
  loc_00DEC1B9: push 00000003h
  loc_00DEC1BB: push 00000000h
  loc_00DEC1BD: call [00401138h] ; __vbaRedim
  loc_00DEC1C3: mov ecx, arg_14
  loc_00DEC1C6: mov edx, arg_10
  loc_00DEC1C9: mov eax, arg_18
  loc_00DEC1CC: add esp, 0000001Ch
  loc_00DEC1CF: mov var_6C, 00000028h
  loc_00DEC1D6: mov var_68, eax
  loc_00DEC1D9: push 00CC0020h
  loc_00DEC1DE: push ecx
  loc_00DEC1DF: mov ecx, arg_C
  loc_00DEC1E2: push edx
  loc_00DEC1E3: mov edx, var_88
  loc_00DEC1E9: push ecx
  loc_00DEC1EA: push edi
  loc_00DEC1EB: push eax
  loc_00DEC1EC: push 00000000h
  loc_00DEC1EE: push 00000000h
  loc_00DEC1F0: push edx
  loc_00DEC1F1: mov var_64, edi
  loc_00DEC1F4: mov var_60, 0001h
  loc_00DEC1FA: mov var_5E, 0018h
  loc_00DEC200: call 006DD820h ; BitBlt(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v)
  loc_00DEC205: call __vbaSetSystemError
  loc_00DEC207: mov eax, var_94
  loc_00DEC20D: mov ecx, arg_18
  loc_00DEC210: mov edx, var_38
  loc_00DEC213: push 00CC0020h
  loc_00DEC218: push 00000000h
  loc_00DEC21A: push 00000000h
  loc_00DEC21C: push eax
  loc_00DEC21D: push edi
  loc_00DEC21E: push ecx
  loc_00DEC21F: push 00000000h
  loc_00DEC221: push 00000000h
  loc_00DEC223: push edx
  loc_00DEC224: call 006DD820h ; BitBlt(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v)
  loc_00DEC229: call __vbaSetSystemError
  loc_00DEC22B: mov eax, var_40
  loc_00DEC22E: lea ecx, var_A8
  loc_00DEC234: push eax
  loc_00DEC235: push ecx
  loc_00DEC236: call [004011DCh] ; __vbaAryLock
  loc_00DEC23C: mov ecx, var_A8
  loc_00DEC242: test ecx, ecx
  loc_00DEC244: jz 00DEC275h
  loc_00DEC246: cmp [ecx], 0001h
  loc_00DEC24A: jnz 00DEC275h
  loc_00DEC24C: mov eax, [ecx+00000014h]
  loc_00DEC24F: mov edx, [ecx+00000010h]
  loc_00DEC252: neg eax
  loc_00DEC254: cmp eax, edx
  loc_00DEC256: mov var_F8, eax
  loc_00DEC25C: jb 00DEC270h
  loc_00DEC25E: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC264: mov ecx, var_A8
  loc_00DEC26A: mov eax, var_F8
  loc_00DEC270: lea eax, [eax+eax*2]
  loc_00DEC273: jmp 00DEC281h
  loc_00DEC275: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC27B: mov ecx, var_A8
  loc_00DEC281: mov ecx, [ecx+0000000Ch]
  loc_00DEC284: lea edx, var_6C
  loc_00DEC287: push 00000000h
  loc_00DEC289: add ecx, eax
  loc_00DEC28B: mov eax, var_88
  loc_00DEC291: push edx
  loc_00DEC292: mov edx, var_7C
  loc_00DEC295: push ecx
  loc_00DEC296: push edi
  loc_00DEC297: push 00000000h
  loc_00DEC299: push edx
  loc_00DEC29A: push eax
  loc_00DEC29B: call 006DD2F4h ; GetDIBits(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v)
  loc_00DEC2A0: call __vbaSetSystemError
  loc_00DEC2A2: lea ecx, var_A8
  loc_00DEC2A8: push ecx
  loc_00DEC2A9: call [00401244h] ; __vbaAryUnlock
  loc_00DEC2AF: mov edx, var_A0
  loc_00DEC2B5: lea eax, var_A8
  loc_00DEC2BB: push edx
  loc_00DEC2BC: push eax
  loc_00DEC2BD: call [004011DCh] ; __vbaAryLock
  loc_00DEC2C3: mov ecx, var_A8
  loc_00DEC2C9: test ecx, ecx
  loc_00DEC2CB: jz 00DEC2FCh
  loc_00DEC2CD: cmp [ecx], 0001h
  loc_00DEC2D1: jnz 00DEC2FCh
  loc_00DEC2D3: mov eax, [ecx+00000014h]
  loc_00DEC2D6: mov edx, [ecx+00000010h]
  loc_00DEC2D9: neg eax
  loc_00DEC2DB: cmp eax, edx
  loc_00DEC2DD: mov var_F8, eax
  loc_00DEC2E3: jb 00DEC2F7h
  loc_00DEC2E5: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC2EB: mov ecx, var_A8
  loc_00DEC2F1: mov eax, var_F8
  loc_00DEC2F7: lea eax, [eax+eax*2]
  loc_00DEC2FA: jmp 00DEC308h
  loc_00DEC2FC: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC302: mov ecx, var_A8
  loc_00DEC308: mov ecx, [ecx+0000000Ch]
  loc_00DEC30B: lea edx, var_6C
  loc_00DEC30E: push 00000000h
  loc_00DEC310: add ecx, eax
  loc_00DEC312: mov eax, var_38
  loc_00DEC315: push edx
  loc_00DEC316: mov edx, var_84
  loc_00DEC31C: push ecx
  loc_00DEC31D: push edi
  loc_00DEC31E: push 00000000h
  loc_00DEC320: push edx
  loc_00DEC321: push eax
  loc_00DEC322: call 006DD2F4h ; GetDIBits(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v)
  loc_00DEC327: call __vbaSetSystemError
  loc_00DEC329: lea ecx, var_A8
  loc_00DEC32F: push ecx
  loc_00DEC330: call [00401244h] ; __vbaAryUnlock
  loc_00DEC336: mov eax, arg_28
  loc_00DEC339: test eax, eax
  loc_00DEC33B: jle 00DEC3A0h
  loc_00DEC33D: cdq
  loc_00DEC33E: and edx, 0000FFFFh
  loc_00DEC344: add eax, edx
  loc_00DEC346: mov ecx, eax
  loc_00DEC348: sar ecx, 10h
  loc_00DEC34B: and ecx, 800000FFh
  loc_00DEC351: jns 00DEC35Bh
  loc_00DEC353: dec ecx
  loc_00DEC354: or ecx, FFFFFF00h
  loc_00DEC35A: inc ecx
  loc_00DEC35B: call [00401158h] ; __vbaUI1I4
  loc_00DEC361: mov var_78, al
  loc_00DEC364: mov eax, arg_28
  loc_00DEC367: cdq
  loc_00DEC368: and edx, 000000FFh
  loc_00DEC36E: add eax, edx
  loc_00DEC370: mov ecx, eax
  loc_00DEC372: sar ecx, 08h
  loc_00DEC375: and ecx, 800000FFh
  loc_00DEC37B: jns 00DEC385h
  loc_00DEC37D: dec ecx
  loc_00DEC37E: or ecx, FFFFFF00h
  loc_00DEC384: inc ecx
  loc_00DEC385: call [00401158h] ; __vbaUI1I4
  loc_00DEC38B: mov ecx, arg_28
  loc_00DEC38E: mov var_77, al
  loc_00DEC391: and ecx, 000000FFh
  loc_00DEC397: call [00401158h] ; __vbaUI1I4
  loc_00DEC39D: mov var_76, al
  loc_00DEC3A0: cmp [ebx+0000009Ch], 0000h
  loc_00DEC3A8: jnz 00DEC3B1h
  loc_00DEC3AA: mov arg_24, FFFFFFFFh
  loc_00DEC3B1: mov edx, arg_18
  loc_00DEC3B4: mov ecx, edi
  loc_00DEC3B6: sub edx, 00000001h
  loc_00DEC3B9: jo 00DED4E3h
  loc_00DEC3BF: sub ecx, 00000001h
  loc_00DEC3C2: mov var_34, edx
  loc_00DEC3C5: jo 00DED4E3h
  loc_00DEC3CB: xor eax, eax
  loc_00DEC3CD: mov var_11C, ecx
  loc_00DEC3D3: mov var_A4, eax
  loc_00DEC3D9: cmp eax, ecx
  loc_00DEC3DB: jg 00DED2FEh
  loc_00DEC3E1: imul eax, arg_18
  loc_00DEC3E5: jo 00DED4E3h
  loc_00DEC3EB: xor ecx, ecx
  loc_00DEC3ED: mov var_98, eax
  loc_00DEC3F3: mov var_74, ecx
  loc_00DEC3F6: cmp ecx, edx
  loc_00DEC3F8: jg 00DED2DAh
  loc_00DEC3FE: add eax, ecx
  loc_00DEC400: jo 00DED4E3h
  loc_00DEC406: mov var_14, eax
  loc_00DEC409: mov eax, [ebx+00000048h]
  loc_00DEC40C: cmp eax, 00000001h
  loc_00DEC40F: jnz 00DEC41Ch
  loc_00DEC411: mov edi, var_18
  loc_00DEC414: and edi, 000000FFh
  loc_00DEC41A: jmp 00DEC495h
  loc_00DEC41C: lea edx, var_90
  loc_00DEC422: mov eax, 00004011h
  loc_00DEC427: mov var_E0, edx
  loc_00DEC42D: lea ecx, [ebx+00000158h]
  loc_00DEC433: mov var_E8, eax
  loc_00DEC439: mov var_D8, eax
  loc_00DEC43F: lea edx, [ebx+0000007Ch]
  loc_00DEC442: mov var_D0, ecx
  loc_00DEC448: lea eax, var_E8
  loc_00DEC44E: mov var_C0, edx
  loc_00DEC454: lea ecx, var_D8
  loc_00DEC45A: push eax
  loc_00DEC45B: lea edx, var_C8
  loc_00DEC461: push ecx
  loc_00DEC462: lea eax, var_B8
  loc_00DEC468: push edx
  loc_00DEC469: push eax
  loc_00DEC46A: mov var_C8, 0000400Bh
  loc_00DEC474: call [004011B4h] ; rtcImmediateIf
  loc_00DEC47A: lea ecx, var_B8
  loc_00DEC480: push ecx
  loc_00DEC481: call [004011D0h] ; __vbaI4Var
  loc_00DEC487: lea ecx, var_B8
  loc_00DEC48D: mov edi, eax
  loc_00DEC48F: call [00401024h] ; __vbaFreeVar
  loc_00DEC495: mov eax, 000000FFh
  loc_00DEC49A: mov edx, [ebx]
  loc_00DEC49C: sub eax, edi
  loc_00DEC49E: jo 00DED4E3h
  loc_00DEC4A4: mov var_8C, eax
  loc_00DEC4AA: lea eax, var_F0
  loc_00DEC4B0: push eax
  loc_00DEC4B1: push ebx
  loc_00DEC4B2: call [edx+000000D8h]
  loc_00DEC4B8: test eax, eax
  loc_00DEC4BA: fnclex
  loc_00DEC4BC: jge 00DEC4D0h
  loc_00DEC4BE: push 000000D8h
  loc_00DEC4C3: push 006DCB00h
  loc_00DEC4C8: push ebx
  loc_00DEC4C9: push eax
  loc_00DEC4CA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEC4D0: mov ecx, var_A0
  loc_00DEC4D6: test ecx, ecx
  loc_00DEC4D8: jz 00DEC506h
  loc_00DEC4DA: cmp [ecx], 0001h
  loc_00DEC4DE: jnz 00DEC506h
  loc_00DEC4E0: mov ebx, var_14
  loc_00DEC4E3: mov edx, [ecx+00000014h]
  loc_00DEC4E6: mov eax, [ecx+00000010h]
  loc_00DEC4E9: sub ebx, edx
  loc_00DEC4EB: cmp ebx, eax
  loc_00DEC4ED: jb 00DEC4FBh
  loc_00DEC4EF: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC4F5: mov ecx, var_A0
  loc_00DEC4FB: lea edx, [ebx+ebx*2]
  loc_00DEC4FE: mov var_14C, edx
  loc_00DEC504: jmp 00DEC518h
  loc_00DEC506: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC50C: mov ecx, var_A0
  loc_00DEC512: mov var_14C, eax
  loc_00DEC518: test ecx, ecx
  loc_00DEC51A: jz 00DEC548h
  loc_00DEC51C: cmp [ecx], 0001h
  loc_00DEC520: jnz 00DEC548h
  loc_00DEC522: mov ebx, var_14
  loc_00DEC525: mov edx, [ecx+00000014h]
  loc_00DEC528: mov eax, [ecx+00000010h]
  loc_00DEC52B: sub ebx, edx
  loc_00DEC52D: cmp ebx, eax
  loc_00DEC52F: jb 00DEC53Dh
  loc_00DEC531: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC537: mov ecx, var_A0
  loc_00DEC53D: lea eax, [ebx+ebx*2]
  loc_00DEC540: mov var_150, eax
  loc_00DEC546: jmp 00DEC55Ah
  loc_00DEC548: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC54E: mov ecx, var_A0
  loc_00DEC554: mov var_150, eax
  loc_00DEC55A: test ecx, ecx
  loc_00DEC55C: jz 00DEC58Ah
  loc_00DEC55E: cmp [ecx], 0001h
  loc_00DEC562: jnz 00DEC58Ah
  loc_00DEC564: mov ebx, var_14
  loc_00DEC567: mov edx, [ecx+00000014h]
  loc_00DEC56A: mov eax, [ecx+00000010h]
  loc_00DEC56D: sub ebx, edx
  loc_00DEC56F: cmp ebx, eax
  loc_00DEC571: jb 00DEC57Fh
  loc_00DEC573: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC579: mov ecx, var_A0
  loc_00DEC57F: lea edx, [ebx+ebx*2]
  loc_00DEC582: mov var_154, edx
  loc_00DEC588: jmp 00DEC59Ch
  loc_00DEC58A: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC590: mov ecx, var_A0
  loc_00DEC596: mov var_154, eax
  loc_00DEC59C: mov ecx, [ecx+0000000Ch]
  loc_00DEC59F: mov edx, var_150
  loc_00DEC5A5: xor eax, eax
  loc_00DEC5A7: mov ebx, var_14C
  loc_00DEC5AD: mov al, [ecx+edx+00000001h]
  loc_00DEC5B1: imul eax, eax, 00000100h
  loc_00DEC5B7: jo 00DED4E3h
  loc_00DEC5BD: xor edx, edx
  loc_00DEC5BF: mov dl, [ecx+ebx+00000002h]
  loc_00DEC5C3: mov ebx, var_154
  loc_00DEC5C9: add eax, edx
  loc_00DEC5CB: jo 00DED4E3h
  loc_00DEC5D1: xor edx, edx
  loc_00DEC5D3: mov dl, [ecx+ebx]
  loc_00DEC5D6: imul edx, edx, 00010000h
  loc_00DEC5DC: jo 00DED4E3h
  loc_00DEC5E2: add eax, edx
  loc_00DEC5E4: jo 00DED4E3h
  loc_00DEC5EA: push eax
  loc_00DEC5EB: mov eax, var_F0
  loc_00DEC5F1: push eax
  loc_00DEC5F2: call 006DD868h ; GetNearestColor(%x1v, %x2v)
  loc_00DEC5F7: mov ebx, eax
  loc_00DEC5F9: call __vbaSetSystemError
  loc_00DEC5FB: mov eax, arg_24
  loc_00DEC5FE: xor ecx, ecx
  loc_00DEC600: cmp ebx, eax
  loc_00DEC602: setnz cl
  loc_00DEC605: neg ecx
  loc_00DEC607: test cx, cx
  loc_00DEC60A: jz 00DED2B3h
  loc_00DEC610: mov edx, var_40
  loc_00DEC613: lea eax, var_108
  loc_00DEC619: push edx
  loc_00DEC61A: push eax
  loc_00DEC61B: call [004011DCh] ; __vbaAryLock
  loc_00DEC621: mov ecx, var_108
  loc_00DEC627: test ecx, ecx
  loc_00DEC629: jz 00DEC653h
  loc_00DEC62B: cmp [ecx], 0001h
  loc_00DEC62F: jnz 00DEC653h
  loc_00DEC631: mov esi, var_14
  loc_00DEC634: mov edx, [ecx+00000014h]
  loc_00DEC637: mov eax, [ecx+00000010h]
  loc_00DEC63A: mov ebx, [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC640: sub esi, edx
  loc_00DEC642: cmp esi, eax
  loc_00DEC644: jb 00DEC64Eh
  loc_00DEC646: call ebx
  loc_00DEC648: mov ecx, var_108
  loc_00DEC64E: lea eax, [esi+esi*2]
  loc_00DEC651: jmp 00DEC661h
  loc_00DEC653: mov ebx, [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC659: call ebx
  loc_00DEC65B: mov ecx, var_108
  loc_00DEC661: mov esi, [ecx+0000000Ch]
  loc_00DEC664: add esi, eax
  loc_00DEC666: mov eax, arg_28
  loc_00DEC669: cmp eax, FFFFFFFFh
  loc_00DEC66C: jle 00DEC876h
  loc_00DEC672: cmp arg_2C, 0000h
  loc_00DEC677: jz 00DEC783h
  loc_00DEC67D: mov ecx, var_A0
  loc_00DEC683: test ecx, ecx
  loc_00DEC685: jz 00DEC6AFh
  loc_00DEC687: cmp [ecx], 0001h
  loc_00DEC68B: jnz 00DEC6AFh
  loc_00DEC68D: mov esi, var_14
  loc_00DEC690: mov edx, [ecx+00000014h]
  loc_00DEC693: mov eax, [ecx+00000010h]
  loc_00DEC696: sub esi, edx
  loc_00DEC698: cmp esi, eax
  loc_00DEC69A: jb 00DEC6A4h
  loc_00DEC69C: call ebx
  loc_00DEC69E: mov ecx, var_A0
  loc_00DEC6A4: lea edx, [esi+esi*2]
  loc_00DEC6A7: mov var_158, edx
  loc_00DEC6AD: jmp 00DEC6BDh
  loc_00DEC6AF: call ebx
  loc_00DEC6B1: mov ecx, var_A0
  loc_00DEC6B7: mov var_158, eax
  loc_00DEC6BD: test ecx, ecx
  loc_00DEC6BF: jz 00DEC6E3h
  loc_00DEC6C1: cmp [ecx], 0001h
  loc_00DEC6C5: jnz 00DEC6E3h
  loc_00DEC6C7: mov esi, var_14
  loc_00DEC6CA: mov edx, [ecx+00000014h]
  loc_00DEC6CD: mov eax, [ecx+00000010h]
  loc_00DEC6D0: sub esi, edx
  loc_00DEC6D2: cmp esi, eax
  loc_00DEC6D4: jb 00DEC6DEh
  loc_00DEC6D6: call ebx
  loc_00DEC6D8: mov ecx, var_A0
  loc_00DEC6DE: lea edi, [esi+esi*2]
  loc_00DEC6E1: jmp 00DEC6EDh
  loc_00DEC6E3: call ebx
  loc_00DEC6E5: mov ecx, var_A0
  loc_00DEC6EB: mov edi, eax
  loc_00DEC6ED: test ecx, ecx
  loc_00DEC6EF: jz 00DEC713h
  loc_00DEC6F1: cmp [ecx], 0001h
  loc_00DEC6F5: jnz 00DEC713h
  loc_00DEC6F7: mov esi, var_14
  loc_00DEC6FA: mov edx, [ecx+00000014h]
  loc_00DEC6FD: mov eax, [ecx+00000010h]
  loc_00DEC700: sub esi, edx
  loc_00DEC702: cmp esi, eax
  loc_00DEC704: jb 00DEC70Eh
  loc_00DEC706: call ebx
  loc_00DEC708: mov ecx, var_A0
  loc_00DEC70E: lea eax, [esi+esi*2]
  loc_00DEC711: jmp 00DEC71Bh
  loc_00DEC713: call ebx
  loc_00DEC715: mov ecx, var_A0
  loc_00DEC71B: mov ecx, [ecx+0000000Ch]
  loc_00DEC71E: mov esi, var_158
  loc_00DEC724: xor edx, edx
  loc_00DEC726: xor ebx, ebx
  loc_00DEC728: mov dl, [ecx+edi+00000001h]
  loc_00DEC72C: mov bl, [ecx+esi+00000002h]
  loc_00DEC730: add edx, ebx
  loc_00DEC732: jo 00DED4E3h
  loc_00DEC738: xor ebx, ebx
  loc_00DEC73A: mov bl, [ecx+eax]
  loc_00DEC73D: add edx, ebx
  loc_00DEC73F: jo 00DED4E3h
  loc_00DEC745: cmp edx, 00000180h
  loc_00DEC74B: jg 00DED292h
  loc_00DEC751: mov ecx, var_40
  loc_00DEC754: test ecx, ecx
  loc_00DEC756: jz 00DEC77Bh
  loc_00DEC758: cmp [ecx], 0001h
  loc_00DEC75C: jnz 00DEC77Bh
  loc_00DEC75E: mov esi, var_14
  loc_00DEC761: mov edx, [ecx+00000014h]
  loc_00DEC764: mov eax, [ecx+00000010h]
  loc_00DEC767: sub esi, edx
  loc_00DEC769: cmp esi, eax
  loc_00DEC76B: jb 00DEC7ACh
  loc_00DEC76D: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC773: mov ecx, var_40
  loc_00DEC776: lea eax, [esi+esi*2]
  loc_00DEC779: jmp 00DEC7B6h
  loc_00DEC77B: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC781: jmp 00DEC7B3h
  loc_00DEC783: cmp edi, 000000FFh
  loc_00DEC789: jnz 00DEC7CDh
  loc_00DEC78B: mov ecx, var_40
  loc_00DEC78E: test ecx, ecx
  loc_00DEC790: jz 00DEC7B1h
  loc_00DEC792: cmp [ecx], 0001h
  loc_00DEC796: jnz 00DEC7B1h
  loc_00DEC798: mov esi, var_14
  loc_00DEC79B: mov edx, [ecx+00000014h]
  loc_00DEC79E: mov eax, [ecx+00000010h]
  loc_00DEC7A1: sub esi, edx
  loc_00DEC7A3: cmp esi, eax
  loc_00DEC7A5: jb 00DEC7ACh
  loc_00DEC7A7: call ebx
  loc_00DEC7A9: mov ecx, var_40
  loc_00DEC7AC: lea eax, [esi+esi*2]
  loc_00DEC7AF: jmp 00DEC7B6h
  loc_00DEC7B1: call ebx
  loc_00DEC7B3: mov ecx, var_40
  loc_00DEC7B6: mov ecx, [ecx+0000000Ch]
  loc_00DEC7B9: lea edx, var_78
  loc_00DEC7BC: add ecx, eax
  loc_00DEC7BE: push edx
  loc_00DEC7BF: push ecx
  loc_00DEC7C0: push 00000003h
  loc_00DEC7C2: call [00401060h] ; __vbaCopyBytes
  loc_00DEC7C8: jmp 00DED292h
  loc_00DEC7CD: test edi, edi
  loc_00DEC7CF: jle 00DED292h
  loc_00DEC7D5: mov eax, var_76
  loc_00DEC7D8: mov ebx, var_8C
  loc_00DEC7DE: and eax, 000000FFh
  loc_00DEC7E3: imul eax, edi
  loc_00DEC7E6: jo 00DED4E3h
  loc_00DEC7EC: xor edx, edx
  loc_00DEC7EE: mov dl, [esi+00000002h]
  loc_00DEC7F1: imul edx, ebx
  loc_00DEC7F4: jo 00DED4E3h
  loc_00DEC7FA: add eax, edx
  loc_00DEC7FC: jo 00DED4E3h
  loc_00DEC802: cdq
  loc_00DEC803: and edx, 000000FFh
  loc_00DEC809: add eax, edx
  loc_00DEC80B: mov ecx, eax
  loc_00DEC80D: sar ecx, 08h
  loc_00DEC810: call [00401158h] ; __vbaUI1I4
  loc_00DEC816: mov [esi+00000002h], al
  loc_00DEC819: xor eax, eax
  loc_00DEC81B: mov al, [esi+00000001h]
  loc_00DEC81E: mov ecx, var_77
  loc_00DEC821: imul eax, ebx
  loc_00DEC824: jo 00DED4E3h
  loc_00DEC82A: and ecx, 000000FFh
  loc_00DEC830: imul ecx, edi
  loc_00DEC833: jo 00DED4E3h
  loc_00DEC839: add eax, ecx
  loc_00DEC83B: jo 00DED4E3h
  loc_00DEC841: cdq
  loc_00DEC842: and edx, 000000FFh
  loc_00DEC848: add eax, edx
  loc_00DEC84A: mov ecx, eax
  loc_00DEC84C: sar ecx, 08h
  loc_00DEC84F: call [00401158h] ; __vbaUI1I4
  loc_00DEC855: mov [esi+00000001h], al
  loc_00DEC858: xor eax, eax
  loc_00DEC85A: mov al, [esi]
  loc_00DEC85C: mov edx, var_78
  loc_00DEC85F: imul eax, ebx
  loc_00DEC862: jo 00DED4E3h
  loc_00DEC868: and edx, 000000FFh
  loc_00DEC86E: imul edx, edi
  loc_00DEC871: jmp 00DED0FFh
  loc_00DEC876: cmp arg_30, 0000h
  loc_00DEC87B: jz 00DECAD2h
  loc_00DEC881: mov eax, var_A0
  loc_00DEC887: test eax, eax
  loc_00DEC889: jz 00DEC8B7h
  loc_00DEC88B: cmp [eax], 0001h
  loc_00DEC88F: jnz 00DEC8B7h
  loc_00DEC891: mov ebx, var_14
  loc_00DEC894: mov edx, [eax+00000014h]
  loc_00DEC897: mov ecx, [eax+00000010h]
  loc_00DEC89A: sub ebx, edx
  loc_00DEC89C: cmp ebx, ecx
  loc_00DEC89E: jb 00DEC8ACh
  loc_00DEC8A0: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC8A6: mov eax, var_A0
  loc_00DEC8AC: lea ecx, [ebx+ebx*2]
  loc_00DEC8AF: mov var_15C, ecx
  loc_00DEC8B5: jmp 00DEC8C5h
  loc_00DEC8B7: call ebx
  loc_00DEC8B9: mov var_15C, eax
  loc_00DEC8BF: mov eax, var_A0
  loc_00DEC8C5: test eax, eax
  loc_00DEC8C7: jz 00DEC8FAh
  loc_00DEC8C9: cmp [eax], 0001h
  loc_00DEC8CD: jnz 00DEC8FAh
  loc_00DEC8CF: mov ecx, var_14
  loc_00DEC8D2: mov edx, [eax+00000014h]
  loc_00DEC8D5: mov ebx, ecx
  loc_00DEC8D7: sub ebx, edx
  loc_00DEC8D9: mov edx, [eax+00000010h]
  loc_00DEC8DC: cmp ebx, edx
  loc_00DEC8DE: jb 00DEC8EFh
  loc_00DEC8E0: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC8E6: mov eax, var_A0
  loc_00DEC8EC: mov ecx, var_14
  loc_00DEC8EF: lea edx, [ebx+ebx*2]
  loc_00DEC8F2: mov var_160, edx
  loc_00DEC8F8: jmp 00DEC90Fh
  loc_00DEC8FA: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC900: mov ecx, var_14
  loc_00DEC903: mov var_160, eax
  loc_00DEC909: mov eax, var_A0
  loc_00DEC90F: test eax, eax
  loc_00DEC911: jz 00DEC93Ch
  loc_00DEC913: cmp [eax], 0001h
  loc_00DEC917: jnz 00DEC93Ch
  loc_00DEC919: sub ecx, [eax+00000014h]
  loc_00DEC91C: mov ebx, ecx
  loc_00DEC91E: mov ecx, [eax+00000010h]
  loc_00DEC921: cmp ebx, ecx
  loc_00DEC923: jb 00DEC931h
  loc_00DEC925: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC92B: mov eax, var_A0
  loc_00DEC931: lea ecx, [ebx+ebx*2]
  loc_00DEC934: mov var_164, ecx
  loc_00DEC93A: jmp 00DEC94Eh
  loc_00DEC93C: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEC942: mov var_164, eax
  loc_00DEC948: mov eax, var_A0
  loc_00DEC94E: mov ebx, [eax+0000000Ch]
  loc_00DEC951: mov ecx, var_15C
  loc_00DEC957: mov edx, ebx
  loc_00DEC959: xor eax, eax
  loc_00DEC95B: mov al, [edx+ecx+00000002h]
  loc_00DEC95F: mov var_168, eax
  loc_00DEC965: fild real4 ptr var_168
  loc_00DEC96B: fstp real8 ptr var_170
  loc_00DEC971: fld real8 ptr var_170
  loc_00DEC977: fmul st0, real8 ptr [00401350h]
  loc_00DEC97D: fnstsw ax
  loc_00DEC97F: test al, 0Dh
  loc_00DEC981: jnz 00DED4DEh
  loc_00DEC987: call [00401208h] ; __vbaFpI4
  loc_00DEC98D: mov var_174, eax
  loc_00DEC993: mov eax, var_160
  loc_00DEC999: fild real4 ptr var_174
  loc_00DEC99F: xor edx, edx
  loc_00DEC9A1: xor ecx, ecx
  loc_00DEC9A3: mov dl, [eax+ebx+00000001h]
  loc_00DEC9A7: fstp real8 ptr var_17C
  loc_00DEC9AD: mov var_180, edx
  loc_00DEC9B3: mov edx, var_164
  loc_00DEC9B9: fild real4 ptr var_180
  loc_00DEC9BF: mov cl, [edx+ebx]
  loc_00DEC9C2: mov var_18C, ecx
  loc_00DEC9C8: fstp real8 ptr var_188
  loc_00DEC9CE: fld real8 ptr var_188
  loc_00DEC9D4: fmul st0, real8 ptr [00401348h]
  loc_00DEC9DA: fadd st0, real8 ptr var_17C
  loc_00DEC9E0: fild real4 ptr var_18C
  loc_00DEC9E6: fstp real8 ptr var_194
  loc_00DEC9EC: fld real8 ptr var_194
  loc_00DEC9F2: fmul st0, real8 ptr [00401340h]
  loc_00DEC9F8: faddp st1
  loc_00DEC9FA: fnstsw ax
  loc_00DEC9FC: test al, 0Dh
  loc_00DEC9FE: jnz 00DED4DEh
  loc_00DECA04: call [00401208h] ; __vbaFpI4
  loc_00DECA0A: cmp edi, 000000FFh
  loc_00DECA10: mov ebx, eax
  loc_00DECA12: jnz 00DECA31h
  loc_00DECA14: mov edi, [00401158h] ; __vbaUI1I4
  loc_00DECA1A: mov ecx, ebx
  loc_00DECA1C: call edi
  loc_00DECA1E: mov ecx, ebx
  loc_00DECA20: mov [esi+00000002h], al
  loc_00DECA23: call edi
  loc_00DECA25: mov ecx, ebx
  loc_00DECA27: mov [esi+00000001h], al
  loc_00DECA2A: call edi
  loc_00DECA2C: jmp 00DED290h
  loc_00DECA31: test edi, edi
  loc_00DECA33: jle 00DED292h
  loc_00DECA39: xor eax, eax
  loc_00DECA3B: mov ecx, edi
  loc_00DECA3D: mov al, [esi+00000002h]
  loc_00DECA40: imul eax, var_8C
  loc_00DECA47: jo 00DED4E3h
  loc_00DECA4D: imul ecx, ebx
  loc_00DECA50: jo 00DED4E3h
  loc_00DECA56: add eax, ecx
  loc_00DECA58: jo 00DED4E3h
  loc_00DECA5E: cdq
  loc_00DECA5F: and edx, 000000FFh
  loc_00DECA65: add eax, edx
  loc_00DECA67: mov ecx, eax
  loc_00DECA69: sar ecx, 08h
  loc_00DECA6C: call [00401158h] ; __vbaUI1I4
  loc_00DECA72: mov [esi+00000002h], al
  loc_00DECA75: xor eax, eax
  loc_00DECA77: mov al, [esi+00000001h]
  loc_00DECA7A: mov edx, edi
  loc_00DECA7C: imul eax, var_8C
  loc_00DECA83: jo 00DED4E3h
  loc_00DECA89: imul edx, ebx
  loc_00DECA8C: jo 00DED4E3h
  loc_00DECA92: add eax, edx
  loc_00DECA94: jo 00DED4E3h
  loc_00DECA9A: cdq
  loc_00DECA9B: and edx, 000000FFh
  loc_00DECAA1: add eax, edx
  loc_00DECAA3: mov ecx, eax
  loc_00DECAA5: sar ecx, 08h
  loc_00DECAA8: call [00401158h] ; __vbaUI1I4
  loc_00DECAAE: mov [esi+00000001h], al
  loc_00DECAB1: xor eax, eax
  loc_00DECAB3: mov al, [esi]
  loc_00DECAB5: imul eax, var_8C
  loc_00DECABC: jo 00DED4E3h
  loc_00DECAC2: imul edi, ebx
  loc_00DECAC5: jo 00DED4E3h
  loc_00DECACB: add eax, edi
  loc_00DECACD: jmp 00DED276h
  loc_00DECAD2: cmp edi, 000000FFh
  loc_00DECAD8: jnz 00DECD9Bh
  loc_00DECADE: mov eax, var_3C
  loc_00DECAE1: cmp eax, 00000001h
  loc_00DECAE4: jnz 00DECBF4h
  loc_00DECAEA: mov ecx, var_A0
  loc_00DECAF0: test ecx, ecx
  loc_00DECAF2: jz 00DECB15h
  loc_00DECAF4: cmp [ecx], ax
  loc_00DECAF7: jnz 00DECB15h
  loc_00DECAF9: mov edi, var_14
  loc_00DECAFC: mov edx, [ecx+00000014h]
  loc_00DECAFF: mov eax, [ecx+00000010h]
  loc_00DECB02: sub edi, edx
  loc_00DECB04: cmp edi, eax
  loc_00DECB06: jb 00DECB10h
  loc_00DECB08: call ebx
  loc_00DECB0A: mov ecx, var_A0
  loc_00DECB10: lea eax, [edi+edi*2]
  loc_00DECB13: jmp 00DECB1Dh
  loc_00DECB15: call ebx
  loc_00DECB17: mov ecx, var_A0
  loc_00DECB1D: mov ecx, [ecx+0000000Ch]
  loc_00DECB20: xor edx, edx
  loc_00DECB22: mov dl, [ecx+eax+00000002h]
  loc_00DECB26: mov edi, edx
  loc_00DECB28: cmp edi, 00000100h
  loc_00DECB2E: jb 00DECB32h
  loc_00DECB30: call ebx
  loc_00DECB32: mov eax, Me
  loc_00DECB35: mov ecx, [eax+0000018Ch]
  loc_00DECB3B: mov dl, [ecx+edi]
  loc_00DECB3E: mov [esi+00000002h], dl
  loc_00DECB41: mov ecx, var_A0
  loc_00DECB47: test ecx, ecx
  loc_00DECB49: jz 00DECB6Dh
  loc_00DECB4B: cmp [ecx], 0001h
  loc_00DECB4F: jnz 00DECB6Dh
  loc_00DECB51: mov edi, var_14
  loc_00DECB54: mov edx, [ecx+00000014h]
  loc_00DECB57: mov eax, [ecx+00000010h]
  loc_00DECB5A: sub edi, edx
  loc_00DECB5C: cmp edi, eax
  loc_00DECB5E: jb 00DECB68h
  loc_00DECB60: call ebx
  loc_00DECB62: mov ecx, var_A0
  loc_00DECB68: lea eax, [edi+edi*2]
  loc_00DECB6B: jmp 00DECB75h
  loc_00DECB6D: call ebx
  loc_00DECB6F: mov ecx, var_A0
  loc_00DECB75: mov ecx, [ecx+0000000Ch]
  loc_00DECB78: xor edx, edx
  loc_00DECB7A: mov dl, [ecx+eax+00000001h]
  loc_00DECB7E: mov edi, edx
  loc_00DECB80: cmp edi, 00000100h
  loc_00DECB86: jb 00DECB8Ah
  loc_00DECB88: call ebx
  loc_00DECB8A: mov eax, Me
  loc_00DECB8D: mov ecx, [eax+0000018Ch]
  loc_00DECB93: mov dl, [ecx+edi]
  loc_00DECB96: mov [esi+00000001h], dl
  loc_00DECB99: mov ecx, var_A0
  loc_00DECB9F: test ecx, ecx
  loc_00DECBA1: jz 00DECBC5h
  loc_00DECBA3: cmp [ecx], 0001h
  loc_00DECBA7: jnz 00DECBC5h
  loc_00DECBA9: mov edi, var_14
  loc_00DECBAC: mov edx, [ecx+00000014h]
  loc_00DECBAF: mov eax, [ecx+00000010h]
  loc_00DECBB2: sub edi, edx
  loc_00DECBB4: cmp edi, eax
  loc_00DECBB6: jb 00DECBC0h
  loc_00DECBB8: call ebx
  loc_00DECBBA: mov ecx, var_A0
  loc_00DECBC0: lea eax, [edi+edi*2]
  loc_00DECBC3: jmp 00DECBCDh
  loc_00DECBC5: call ebx
  loc_00DECBC7: mov ecx, var_A0
  loc_00DECBCD: mov ecx, [ecx+0000000Ch]
  loc_00DECBD0: xor edx, edx
  loc_00DECBD2: mov dl, [ecx+eax]
  loc_00DECBD5: mov edi, edx
  loc_00DECBD7: cmp edi, 00000100h
  loc_00DECBDD: jb 00DECBE1h
  loc_00DECBDF: call ebx
  loc_00DECBE1: mov eax, Me
  loc_00DECBE4: mov ecx, [eax+0000018Ch]
  loc_00DECBEA: mov dl, [ecx+edi]
  loc_00DECBED: mov [esi], dl
  loc_00DECBEF: jmp 00DED292h
  loc_00DECBF4: cmp eax, 00000002h
  loc_00DECBF7: jnz 00DECD08h
  loc_00DECBFD: mov ecx, var_A0
  loc_00DECC03: test ecx, ecx
  loc_00DECC05: jz 00DECC29h
  loc_00DECC07: cmp [ecx], 0001h
  loc_00DECC0B: jnz 00DECC29h
  loc_00DECC0D: mov edi, var_14
  loc_00DECC10: mov edx, [ecx+00000014h]
  loc_00DECC13: mov eax, [ecx+00000010h]
  loc_00DECC16: sub edi, edx
  loc_00DECC18: cmp edi, eax
  loc_00DECC1A: jb 00DECC24h
  loc_00DECC1C: call ebx
  loc_00DECC1E: mov ecx, var_A0
  loc_00DECC24: lea eax, [edi+edi*2]
  loc_00DECC27: jmp 00DECC31h
  loc_00DECC29: call ebx
  loc_00DECC2B: mov ecx, var_A0
  loc_00DECC31: mov ecx, [ecx+0000000Ch]
  loc_00DECC34: xor edx, edx
  loc_00DECC36: mov dl, [ecx+eax+00000002h]
  loc_00DECC3A: mov edi, edx
  loc_00DECC3C: cmp edi, 00000100h
  loc_00DECC42: jb 00DECC46h
  loc_00DECC44: call ebx
  loc_00DECC46: mov eax, Me
  loc_00DECC49: mov ecx, [eax+000001A8h]
  loc_00DECC4F: mov dl, [ecx+edi]
  loc_00DECC52: mov [esi+00000002h], dl
  loc_00DECC55: mov ecx, var_A0
  loc_00DECC5B: test ecx, ecx
  loc_00DECC5D: jz 00DECC81h
  loc_00DECC5F: cmp [ecx], 0001h
  loc_00DECC63: jnz 00DECC81h
  loc_00DECC65: mov edi, var_14
  loc_00DECC68: mov edx, [ecx+00000014h]
  loc_00DECC6B: mov eax, [ecx+00000010h]
  loc_00DECC6E: sub edi, edx
  loc_00DECC70: cmp edi, eax
  loc_00DECC72: jb 00DECC7Ch
  loc_00DECC74: call ebx
  loc_00DECC76: mov ecx, var_A0
  loc_00DECC7C: lea eax, [edi+edi*2]
  loc_00DECC7F: jmp 00DECC89h
  loc_00DECC81: call ebx
  loc_00DECC83: mov ecx, var_A0
  loc_00DECC89: mov ecx, [ecx+0000000Ch]
  loc_00DECC8C: xor edx, edx
  loc_00DECC8E: mov dl, [ecx+eax+00000001h]
  loc_00DECC92: mov edi, edx
  loc_00DECC94: cmp edi, 00000100h
  loc_00DECC9A: jb 00DECC9Eh
  loc_00DECC9C: call ebx
  loc_00DECC9E: mov eax, Me
  loc_00DECCA1: mov ecx, [eax+000001A8h]
  loc_00DECCA7: mov dl, [ecx+edi]
  loc_00DECCAA: mov [esi+00000001h], dl
  loc_00DECCAD: mov ecx, var_A0
  loc_00DECCB3: test ecx, ecx
  loc_00DECCB5: jz 00DECCD9h
  loc_00DECCB7: cmp [ecx], 0001h
  loc_00DECCBB: jnz 00DECCD9h
  loc_00DECCBD: mov edi, var_14
  loc_00DECCC0: mov edx, [ecx+00000014h]
  loc_00DECCC3: mov eax, [ecx+00000010h]
  loc_00DECCC6: sub edi, edx
  loc_00DECCC8: cmp edi, eax
  loc_00DECCCA: jb 00DECCD4h
  loc_00DECCCC: call ebx
  loc_00DECCCE: mov ecx, var_A0
  loc_00DECCD4: lea eax, [edi+edi*2]
  loc_00DECCD7: jmp 00DECCE1h
  loc_00DECCD9: call ebx
  loc_00DECCDB: mov ecx, var_A0
  loc_00DECCE1: mov ecx, [ecx+0000000Ch]
  loc_00DECCE4: xor edx, edx
  loc_00DECCE6: mov dl, [ecx+eax]
  loc_00DECCE9: mov edi, edx
  loc_00DECCEB: cmp edi, 00000100h
  loc_00DECCF1: jb 00DECCF5h
  loc_00DECCF3: call ebx
  loc_00DECCF5: mov eax, Me
  loc_00DECCF8: mov ecx, [eax+000001A8h]
  loc_00DECCFE: mov dl, [ecx+edi]
  loc_00DECD01: mov [esi], dl
  loc_00DECD03: jmp 00DED292h
  loc_00DECD08: mov eax, var_A0
  loc_00DECD0E: test eax, eax
  loc_00DECD10: jz 00DECD30h
  loc_00DECD12: cmp [eax], 0001h
  loc_00DECD16: jnz 00DECD30h
  loc_00DECD18: mov edi, var_14
  loc_00DECD1B: mov edx, [eax+00000014h]
  loc_00DECD1E: mov ecx, [eax+00000010h]
  loc_00DECD21: mov esi, edi
  loc_00DECD23: sub esi, edx
  loc_00DECD25: cmp esi, ecx
  loc_00DECD27: jb 00DECD2Bh
  loc_00DECD29: call ebx
  loc_00DECD2B: lea esi, [esi+esi*2]
  loc_00DECD2E: jmp 00DECD37h
  loc_00DECD30: call ebx
  loc_00DECD32: mov edi, var_14
  loc_00DECD35: mov esi, eax
  loc_00DECD37: mov ecx, var_40
  loc_00DECD3A: test ecx, ecx
  loc_00DECD3C: jz 00DECD77h
  loc_00DECD3E: cmp [ecx], 0001h
  loc_00DECD42: jnz 00DECD77h
  loc_00DECD44: mov edx, [ecx+00000014h]
  loc_00DECD47: mov eax, [ecx+00000010h]
  loc_00DECD4A: sub edi, edx
  loc_00DECD4C: cmp edi, eax
  loc_00DECD4E: jb 00DECD55h
  loc_00DECD50: call ebx
  loc_00DECD52: mov ecx, var_40
  loc_00DECD55: mov edx, var_A0
  loc_00DECD5B: mov ecx, [ecx+0000000Ch]
  loc_00DECD5E: lea eax, [edi+edi*2]
  loc_00DECD61: mov edx, [edx+0000000Ch]
  loc_00DECD64: add ecx, eax
  loc_00DECD66: add edx, esi
  loc_00DECD68: push edx
  loc_00DECD69: push ecx
  loc_00DECD6A: push 00000003h
  loc_00DECD6C: call [00401060h] ; __vbaCopyBytes
  loc_00DECD72: jmp 00DED292h
  loc_00DECD77: call ebx
  loc_00DECD79: mov edx, var_A0
  loc_00DECD7F: mov ecx, var_40
  loc_00DECD82: mov edx, [edx+0000000Ch]
  loc_00DECD85: mov ecx, [ecx+0000000Ch]
  loc_00DECD88: add edx, esi
  loc_00DECD8A: add ecx, eax
  loc_00DECD8C: push edx
  loc_00DECD8D: push ecx
  loc_00DECD8E: push 00000003h
  loc_00DECD90: call [00401060h] ; __vbaCopyBytes
  loc_00DECD96: jmp 00DED292h
  loc_00DECD9B: test edi, edi
  loc_00DECD9D: jle 00DED292h
  loc_00DECDA3: mov eax, var_3C
  loc_00DECDA6: cmp eax, 00000001h
  loc_00DECDA9: jnz 00DECF46h
  loc_00DECDAF: mov ecx, var_A0
  loc_00DECDB5: test ecx, ecx
  loc_00DECDB7: jz 00DECDDEh
  loc_00DECDB9: cmp [ecx], ax
  loc_00DECDBC: jnz 00DECDDEh
  loc_00DECDBE: mov ebx, var_14
  loc_00DECDC1: mov edx, [ecx+00000014h]
  loc_00DECDC4: mov eax, [ecx+00000010h]
  loc_00DECDC7: sub ebx, edx
  loc_00DECDC9: cmp ebx, eax
  loc_00DECDCB: jb 00DECDD9h
  loc_00DECDCD: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECDD3: mov ecx, var_A0
  loc_00DECDD9: lea eax, [ebx+ebx*2]
  loc_00DECDDC: jmp 00DECDEAh
  loc_00DECDDE: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECDE4: mov ecx, var_A0
  loc_00DECDEA: mov edx, [ecx+0000000Ch]
  loc_00DECDED: xor ebx, ebx
  loc_00DECDEF: mov bl, [edx+eax+00000002h]
  loc_00DECDF3: cmp ebx, 00000100h
  loc_00DECDF9: jb 00DECE01h
  loc_00DECDFB: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECE01: mov eax, Me
  loc_00DECE04: mov ecx, [eax+0000018Ch]
  loc_00DECE0A: xor eax, eax
  loc_00DECE0C: mov al, [ecx+ebx]
  loc_00DECE0F: imul eax, edi
  loc_00DECE12: jo 00DED4E3h
  loc_00DECE18: xor edx, edx
  loc_00DECE1A: mov dl, [esi+00000002h]
  loc_00DECE1D: imul edx, var_8C
  loc_00DECE24: jo 00DED4E3h
  loc_00DECE2A: add eax, edx
  loc_00DECE2C: jo 00DED4E3h
  loc_00DECE32: cdq
  loc_00DECE33: and edx, 000000FFh
  loc_00DECE39: add eax, edx
  loc_00DECE3B: mov ecx, eax
  loc_00DECE3D: sar ecx, 08h
  loc_00DECE40: call [00401158h] ; __vbaUI1I4
  loc_00DECE46: mov [esi+00000002h], al
  loc_00DECE49: mov ecx, var_A0
  loc_00DECE4F: test ecx, ecx
  loc_00DECE51: jz 00DECE79h
  loc_00DECE53: cmp [ecx], 0001h
  loc_00DECE57: jnz 00DECE79h
  loc_00DECE59: mov ebx, var_14
  loc_00DECE5C: mov edx, [ecx+00000014h]
  loc_00DECE5F: mov eax, [ecx+00000010h]
  loc_00DECE62: sub ebx, edx
  loc_00DECE64: cmp ebx, eax
  loc_00DECE66: jb 00DECE74h
  loc_00DECE68: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECE6E: mov ecx, var_A0
  loc_00DECE74: lea eax, [ebx+ebx*2]
  loc_00DECE77: jmp 00DECE85h
  loc_00DECE79: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECE7F: mov ecx, var_A0
  loc_00DECE85: mov ecx, [ecx+0000000Ch]
  loc_00DECE88: xor ebx, ebx
  loc_00DECE8A: mov bl, [ecx+eax+00000001h]
  loc_00DECE8E: cmp ebx, 00000100h
  loc_00DECE94: jb 00DECE9Ch
  loc_00DECE96: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECE9C: mov edx, Me
  loc_00DECE9F: xor ecx, ecx
  loc_00DECEA1: mov eax, [edx+0000018Ch]
  loc_00DECEA7: mov cl, [eax+ebx]
  loc_00DECEAA: mov eax, ecx
  loc_00DECEAC: imul eax, edi
  loc_00DECEAF: jo 00DED4E3h
  loc_00DECEB5: xor edx, edx
  loc_00DECEB7: mov dl, [esi+00000001h]
  loc_00DECEBA: imul edx, var_8C
  loc_00DECEC1: jo 00DED4E3h
  loc_00DECEC7: add eax, edx
  loc_00DECEC9: jo 00DED4E3h
  loc_00DECECF: cdq
  loc_00DECED0: and edx, 000000FFh
  loc_00DECED6: add eax, edx
  loc_00DECED8: mov ecx, eax
  loc_00DECEDA: sar ecx, 08h
  loc_00DECEDD: call [00401158h] ; __vbaUI1I4
  loc_00DECEE3: mov [esi+00000001h], al
  loc_00DECEE6: mov ecx, var_A0
  loc_00DECEEC: test ecx, ecx
  loc_00DECEEE: jz 00DECF16h
  loc_00DECEF0: cmp [ecx], 0001h
  loc_00DECEF4: jnz 00DECF16h
  loc_00DECEF6: mov ebx, var_14
  loc_00DECEF9: mov edx, [ecx+00000014h]
  loc_00DECEFC: mov eax, [ecx+00000010h]
  loc_00DECEFF: sub ebx, edx
  loc_00DECF01: cmp ebx, eax
  loc_00DECF03: jb 00DECF11h
  loc_00DECF05: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECF0B: mov ecx, var_A0
  loc_00DECF11: lea eax, [ebx+ebx*2]
  loc_00DECF14: jmp 00DECF22h
  loc_00DECF16: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECF1C: mov ecx, var_A0
  loc_00DECF22: mov ecx, [ecx+0000000Ch]
  loc_00DECF25: xor ebx, ebx
  loc_00DECF27: mov bl, [ecx+eax]
  loc_00DECF2A: cmp ebx, 00000100h
  loc_00DECF30: jb 00DECF38h
  loc_00DECF32: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECF38: mov edx, Me
  loc_00DECF3B: mov eax, [edx+0000018Ch]
  loc_00DECF41: jmp 00DED0E4h
  loc_00DECF46: mov ecx, var_A0
  loc_00DECF4C: cmp eax, 00000002h
  loc_00DECF4F: jnz 00DED10Ch
  loc_00DECF55: test ecx, ecx
  loc_00DECF57: jz 00DECF7Fh
  loc_00DECF59: cmp [ecx], 0001h
  loc_00DECF5D: jnz 00DECF7Fh
  loc_00DECF5F: mov ebx, var_14
  loc_00DECF62: mov edx, [ecx+00000014h]
  loc_00DECF65: mov eax, [ecx+00000010h]
  loc_00DECF68: sub ebx, edx
  loc_00DECF6A: cmp ebx, eax
  loc_00DECF6C: jb 00DECF7Ah
  loc_00DECF6E: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECF74: mov ecx, var_A0
  loc_00DECF7A: lea eax, [ebx+ebx*2]
  loc_00DECF7D: jmp 00DECF8Bh
  loc_00DECF7F: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECF85: mov ecx, var_A0
  loc_00DECF8B: mov ecx, [ecx+0000000Ch]
  loc_00DECF8E: xor ebx, ebx
  loc_00DECF90: mov bl, [ecx+eax+00000002h]
  loc_00DECF94: cmp ebx, 00000100h
  loc_00DECF9A: jb 00DECFA2h
  loc_00DECF9C: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DECFA2: mov edx, Me
  loc_00DECFA5: xor ecx, ecx
  loc_00DECFA7: mov eax, [edx+000001A8h]
  loc_00DECFAD: mov cl, [eax+ebx]
  loc_00DECFB0: mov eax, ecx
  loc_00DECFB2: imul eax, edi
  loc_00DECFB5: jo 00DED4E3h
  loc_00DECFBB: xor edx, edx
  loc_00DECFBD: mov dl, [esi+00000002h]
  loc_00DECFC0: imul edx, var_8C
  loc_00DECFC7: jo 00DED4E3h
  loc_00DECFCD: add eax, edx
  loc_00DECFCF: jo 00DED4E3h
  loc_00DECFD5: cdq
  loc_00DECFD6: and edx, 000000FFh
  loc_00DECFDC: add eax, edx
  loc_00DECFDE: mov ecx, eax
  loc_00DECFE0: sar ecx, 08h
  loc_00DECFE3: call [00401158h] ; __vbaUI1I4
  loc_00DECFE9: mov [esi+00000002h], al
  loc_00DECFEC: mov ecx, var_A0
  loc_00DECFF2: test ecx, ecx
  loc_00DECFF4: jz 00DED01Ch
  loc_00DECFF6: cmp [ecx], 0001h
  loc_00DECFFA: jnz 00DED01Ch
  loc_00DECFFC: mov ebx, var_14
  loc_00DECFFF: mov edx, [ecx+00000014h]
  loc_00DED002: mov eax, [ecx+00000010h]
  loc_00DED005: sub ebx, edx
  loc_00DED007: cmp ebx, eax
  loc_00DED009: jb 00DED017h
  loc_00DED00B: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED011: mov ecx, var_A0
  loc_00DED017: lea eax, [ebx+ebx*2]
  loc_00DED01A: jmp 00DED028h
  loc_00DED01C: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED022: mov ecx, var_A0
  loc_00DED028: mov ecx, [ecx+0000000Ch]
  loc_00DED02B: xor ebx, ebx
  loc_00DED02D: mov bl, [ecx+eax+00000001h]
  loc_00DED031: cmp ebx, 00000100h
  loc_00DED037: jb 00DED03Fh
  loc_00DED039: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED03F: mov edx, Me
  loc_00DED042: xor ecx, ecx
  loc_00DED044: mov eax, [edx+000001A8h]
  loc_00DED04A: mov cl, [eax+ebx]
  loc_00DED04D: mov eax, ecx
  loc_00DED04F: imul eax, edi
  loc_00DED052: jo 00DED4E3h
  loc_00DED058: xor edx, edx
  loc_00DED05A: mov dl, [esi+00000001h]
  loc_00DED05D: imul edx, var_8C
  loc_00DED064: jo 00DED4E3h
  loc_00DED06A: add eax, edx
  loc_00DED06C: jo 00DED4E3h
  loc_00DED072: cdq
  loc_00DED073: and edx, 000000FFh
  loc_00DED079: add eax, edx
  loc_00DED07B: mov ecx, eax
  loc_00DED07D: sar ecx, 08h
  loc_00DED080: call [00401158h] ; __vbaUI1I4
  loc_00DED086: mov [esi+00000001h], al
  loc_00DED089: mov ecx, var_A0
  loc_00DED08F: test ecx, ecx
  loc_00DED091: jz 00DED0B9h
  loc_00DED093: cmp [ecx], 0001h
  loc_00DED097: jnz 00DED0B9h
  loc_00DED099: mov ebx, var_14
  loc_00DED09C: mov edx, [ecx+00000014h]
  loc_00DED09F: mov eax, [ecx+00000010h]
  loc_00DED0A2: sub ebx, edx
  loc_00DED0A4: cmp ebx, eax
  loc_00DED0A6: jb 00DED0B4h
  loc_00DED0A8: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED0AE: mov ecx, var_A0
  loc_00DED0B4: lea eax, [ebx+ebx*2]
  loc_00DED0B7: jmp 00DED0C5h
  loc_00DED0B9: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED0BF: mov ecx, var_A0
  loc_00DED0C5: mov ecx, [ecx+0000000Ch]
  loc_00DED0C8: xor ebx, ebx
  loc_00DED0CA: mov bl, [ecx+eax]
  loc_00DED0CD: cmp ebx, 00000100h
  loc_00DED0D3: jb 00DED0DBh
  loc_00DED0D5: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED0DB: mov edx, Me
  loc_00DED0DE: mov eax, [edx+000001A8h]
  loc_00DED0E4: xor ecx, ecx
  loc_00DED0E6: mov cl, [eax+ebx]
  loc_00DED0E9: mov eax, ecx
  loc_00DED0EB: imul eax, edi
  loc_00DED0EE: jo 00DED4E3h
  loc_00DED0F4: xor edx, edx
  loc_00DED0F6: mov dl, [esi]
  loc_00DED0F8: imul edx, var_8C
  loc_00DED0FF: jo 00DED4E3h
  loc_00DED105: add eax, edx
  loc_00DED107: jmp 00DED276h
  loc_00DED10C: test ecx, ecx
  loc_00DED10E: jz 00DED136h
  loc_00DED110: cmp [ecx], 0001h
  loc_00DED114: jnz 00DED136h
  loc_00DED116: mov ebx, var_14
  loc_00DED119: mov edx, [ecx+00000014h]
  loc_00DED11C: mov eax, [ecx+00000010h]
  loc_00DED11F: sub ebx, edx
  loc_00DED121: cmp ebx, eax
  loc_00DED123: jb 00DED131h
  loc_00DED125: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED12B: mov ecx, var_A0
  loc_00DED131: lea eax, [ebx+ebx*2]
  loc_00DED134: jmp 00DED142h
  loc_00DED136: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED13C: mov ecx, var_A0
  loc_00DED142: mov ecx, [ecx+0000000Ch]
  loc_00DED145: xor edx, edx
  loc_00DED147: mov ebx, var_8C
  loc_00DED14D: mov dl, [ecx+eax+00000002h]
  loc_00DED151: mov eax, edx
  loc_00DED153: imul eax, edi
  loc_00DED156: jo 00DED4E3h
  loc_00DED15C: xor ecx, ecx
  loc_00DED15E: mov cl, [esi+00000002h]
  loc_00DED161: imul ecx, ebx
  loc_00DED164: jo 00DED4E3h
  loc_00DED16A: add eax, ecx
  loc_00DED16C: jo 00DED4E3h
  loc_00DED172: cdq
  loc_00DED173: and edx, 000000FFh
  loc_00DED179: add eax, edx
  loc_00DED17B: mov ecx, eax
  loc_00DED17D: sar ecx, 08h
  loc_00DED180: call [00401158h] ; __vbaUI1I4
  loc_00DED186: mov [esi+00000002h], al
  loc_00DED189: mov ecx, var_A0
  loc_00DED18F: test ecx, ecx
  loc_00DED191: jz 00DED1BFh
  loc_00DED193: cmp [ecx], 0001h
  loc_00DED197: jnz 00DED1BFh
  loc_00DED199: mov ebx, var_14
  loc_00DED19C: mov edx, [ecx+00000014h]
  loc_00DED19F: mov eax, [ecx+00000010h]
  loc_00DED1A2: sub ebx, edx
  loc_00DED1A4: cmp ebx, eax
  loc_00DED1A6: jb 00DED1B4h
  loc_00DED1A8: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED1AE: mov ecx, var_A0
  loc_00DED1B4: lea eax, [ebx+ebx*2]
  loc_00DED1B7: mov ebx, var_8C
  loc_00DED1BD: jmp 00DED1CBh
  loc_00DED1BF: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED1C5: mov ecx, var_A0
  loc_00DED1CB: mov edx, [ecx+0000000Ch]
  loc_00DED1CE: xor ecx, ecx
  loc_00DED1D0: mov cl, [edx+eax+00000001h]
  loc_00DED1D4: mov eax, ecx
  loc_00DED1D6: imul eax, edi
  loc_00DED1D9: jo 00DED4E3h
  loc_00DED1DF: xor edx, edx
  loc_00DED1E1: mov dl, [esi+00000001h]
  loc_00DED1E4: imul edx, ebx
  loc_00DED1E7: jo 00DED4E3h
  loc_00DED1ED: add eax, edx
  loc_00DED1EF: jo 00DED4E3h
  loc_00DED1F5: cdq
  loc_00DED1F6: and edx, 000000FFh
  loc_00DED1FC: add eax, edx
  loc_00DED1FE: mov ecx, eax
  loc_00DED200: sar ecx, 08h
  loc_00DED203: call [00401158h] ; __vbaUI1I4
  loc_00DED209: mov [esi+00000001h], al
  loc_00DED20C: mov ecx, var_A0
  loc_00DED212: test ecx, ecx
  loc_00DED214: jz 00DED248h
  loc_00DED216: cmp [ecx], 0001h
  loc_00DED21A: jnz 00DED248h
  loc_00DED21C: mov eax, var_14
  loc_00DED21F: mov edx, [ecx+00000014h]
  loc_00DED222: sub eax, edx
  loc_00DED224: mov edx, [ecx+00000010h]
  loc_00DED227: cmp eax, edx
  loc_00DED229: mov var_F8, eax
  loc_00DED22F: jb 00DED243h
  loc_00DED231: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED237: mov ecx, var_A0
  loc_00DED23D: mov eax, var_F8
  loc_00DED243: lea eax, [eax+eax*2]
  loc_00DED246: jmp 00DED254h
  loc_00DED248: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED24E: mov ecx, var_A0
  loc_00DED254: mov ecx, [ecx+0000000Ch]
  loc_00DED257: xor edx, edx
  loc_00DED259: mov dl, [ecx+eax]
  loc_00DED25C: mov eax, edx
  loc_00DED25E: imul eax, edi
  loc_00DED261: jo 00DED4E3h
  loc_00DED267: xor ecx, ecx
  loc_00DED269: mov cl, [esi]
  loc_00DED26B: imul ecx, ebx
  loc_00DED26E: jo 00DED4E3h
  loc_00DED274: add eax, ecx
  loc_00DED276: jo 00DED4E3h
  loc_00DED27C: cdq
  loc_00DED27D: and edx, 000000FFh
  loc_00DED283: add eax, edx
  loc_00DED285: mov ecx, eax
  loc_00DED287: sar ecx, 08h
  loc_00DED28A: call [00401158h] ; __vbaUI1I4
  loc_00DED290: mov [esi], al
  loc_00DED292: lea edx, var_108
  loc_00DED298: push edx
  loc_00DED299: call [00401244h] ; __vbaAryUnlock
  loc_00DED29F: mov esi, [00401074h] ; __vbaSetSystemError
  loc_00DED2A5: mov eax, var_98
  loc_00DED2AB: mov ebx, Me
  loc_00DED2AE: mov edi, arg_1C
  loc_00DED2B1: jmp 00DED2BFh
  loc_00DED2B3: mov edi, arg_1C
  loc_00DED2B6: mov ebx, Me
  loc_00DED2B9: mov eax, var_98
  loc_00DED2BF: mov edx, var_74
  loc_00DED2C2: mov ecx, 00000001h
  loc_00DED2C7: add ecx, edx
  loc_00DED2C9: mov edx, var_34
  loc_00DED2CC: jo 00DED4E3h
  loc_00DED2D2: mov var_74, ecx
  loc_00DED2D5: jmp 00DEC3F6h
  loc_00DED2DA: mov ecx, var_A4
  loc_00DED2E0: mov eax, 00000001h
  loc_00DED2E5: add eax, ecx
  loc_00DED2E7: mov ecx, var_11C
  loc_00DED2ED: jo 00DED4E3h
  loc_00DED2F3: mov var_A4, eax
  loc_00DED2F9: jmp 00DEC3D9h
  loc_00DED2FE: mov eax, var_40
  loc_00DED301: lea ecx, var_A8
  loc_00DED307: push eax
  loc_00DED308: push ecx
  loc_00DED309: call [004011DCh] ; __vbaAryLock
  loc_00DED30F: mov ecx, var_A8
  loc_00DED315: test ecx, ecx
  loc_00DED317: jz 00DED33Ch
  loc_00DED319: cmp [ecx], 0001h
  loc_00DED31D: jnz 00DED33Ch
  loc_00DED31F: mov ebx, [ecx+00000014h]
  loc_00DED322: mov eax, [ecx+00000010h]
  loc_00DED325: neg ebx
  loc_00DED327: cmp ebx, eax
  loc_00DED329: jb 00DED337h
  loc_00DED32B: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED331: mov ecx, var_A8
  loc_00DED337: lea eax, [ebx+ebx*2]
  loc_00DED33A: jmp 00DED348h
  loc_00DED33C: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DED342: mov ecx, var_A8
  loc_00DED348: mov ecx, [ecx+0000000Ch]
  loc_00DED34B: lea edx, var_6C
  loc_00DED34E: push 00000000h
  loc_00DED350: add ecx, eax
  loc_00DED352: mov eax, arg_14
  loc_00DED355: push edx
  loc_00DED356: mov edx, arg_18
  loc_00DED359: push ecx
  loc_00DED35A: mov ecx, arg_10
  loc_00DED35D: push edi
  loc_00DED35E: push 00000000h
  loc_00DED360: push 00000000h
  loc_00DED362: push 00000000h
  loc_00DED364: push edi
  loc_00DED365: push edx
  loc_00DED366: mov edx, arg_C
  loc_00DED369: push eax
  loc_00DED36A: push ecx
  loc_00DED36B: push edx
  loc_00DED36C: call 006DD340h ; SetDIBitsToDevice(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v, %x10v, %x11v, %x12v)
  loc_00DED371: call __vbaSetSystemError
  loc_00DED373: lea eax, var_A8
  loc_00DED379: push eax
  loc_00DED37A: call [00401244h] ; __vbaAryUnlock
  loc_00DED380: mov edi, [004010E8h] ; __vbaErase
  loc_00DED386: lea ecx, var_A0
  loc_00DED38C: push ecx
  loc_00DED38D: push 00000000h
  loc_00DED38F: call edi
  loc_00DED391: lea edx, var_40
  loc_00DED394: push edx
  loc_00DED395: push 00000000h
  loc_00DED397: call edi
  loc_00DED399: mov eax, var_20
  loc_00DED39C: mov ebx, var_88
  loc_00DED3A2: push eax
  loc_00DED3A3: push ebx
  loc_00DED3A4: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DED3A9: mov var_F0, eax
  loc_00DED3AF: call __vbaSetSystemError
  loc_00DED3B1: mov ecx, var_F0
  loc_00DED3B7: push ecx
  loc_00DED3B8: call 006DD498h ; DeleteObject(%x1v)
  loc_00DED3BD: call __vbaSetSystemError
  loc_00DED3BF: mov edx, var_30
  loc_00DED3C2: mov eax, var_38
  loc_00DED3C5: push edx
  loc_00DED3C6: push eax
  loc_00DED3C7: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DED3CC: mov var_F0, eax
  loc_00DED3D2: call __vbaSetSystemError
  loc_00DED3D4: mov ecx, var_F0
  loc_00DED3DA: push ecx
  loc_00DED3DB: call 006DD498h ; DeleteObject(%x1v)
  loc_00DED3E0: call __vbaSetSystemError
  loc_00DED3E2: mov edx, var_24
  loc_00DED3E5: push 00000000h
  loc_00DED3E7: push 00000003h
  loc_00DED3E9: lea eax, var_B8
  loc_00DED3EF: push edx
  loc_00DED3F0: push eax
  loc_00DED3F1: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DED3F7: add esp, 00000010h
  loc_00DED3FA: push eax
  loc_00DED3FB: call [0040118Ch] ; __vbaI2Var
  loc_00DED401: xor ecx, ecx
  loc_00DED403: cmp ax, 0003h
  loc_00DED407: setz cl
  loc_00DED40A: neg ecx
  loc_00DED40C: mov edi, ecx
  loc_00DED40E: lea ecx, var_B8
  loc_00DED414: call [00401024h] ; __vbaFreeVar
  loc_00DED41A: test di, di
  loc_00DED41D: jz 00DED448h
  loc_00DED41F: mov edx, var_9C
  loc_00DED425: mov eax, var_94
  loc_00DED42B: push edx
  loc_00DED42C: push eax
  loc_00DED42D: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DED432: mov var_F0, eax
  loc_00DED438: call __vbaSetSystemError
  loc_00DED43A: mov ecx, var_F0
  loc_00DED440: push ecx
  loc_00DED441: call 006DD498h ; DeleteObject(%x1v)
  loc_00DED446: call __vbaSetSystemError
  loc_00DED448: push ebx
  loc_00DED449: call 006DD4DCh ; DeleteDC(%x1v)
  loc_00DED44E: call __vbaSetSystemError
  loc_00DED450: mov edx, var_38
  loc_00DED453: push edx
  loc_00DED454: call 006DD4DCh ; DeleteDC(%x1v)
  loc_00DED459: call __vbaSetSystemError
  loc_00DED45B: mov eax, var_9C
  loc_00DED461: push eax
  loc_00DED462: call 006DD498h ; DeleteObject(%x1v)
  loc_00DED467: call __vbaSetSystemError
  loc_00DED469: mov ecx, var_94
  loc_00DED46F: push ecx
  loc_00DED470: call 006DD4DCh ; DeleteDC(%x1v)
  loc_00DED475: call __vbaSetSystemError
  loc_00DED477: fwait
  loc_00DED478: push 00DED4C9h
  loc_00DED47D: jmp 00DED499h
  loc_00DED47F: lea edx, var_A8
  loc_00DED485: push edx
  loc_00DED486: call [00401244h] ; __vbaAryUnlock
  loc_00DED48C: lea ecx, var_B8
  loc_00DED492: call [00401024h] ; __vbaFreeVar
  loc_00DED498: ret
  loc_00DED499: lea eax, var_108
  loc_00DED49F: push eax
  loc_00DED4A0: call [00401244h] ; __vbaAryUnlock
  loc_00DED4A6: lea ecx, var_24
  loc_00DED4A9: call [00401254h] ; __vbaFreeObj
  loc_00DED4AF: mov esi, [00401090h] ; __vbaAryDestruct
  loc_00DED4B5: lea ecx, var_40
  loc_00DED4B8: push ecx
  loc_00DED4B9: push 00000000h
  loc_00DED4BB: call __vbaAryDestruct
  loc_00DED4BD: lea edx, var_A0
  loc_00DED4C3: push edx
  loc_00DED4C4: push 00000000h
  loc_00DED4C6: call __vbaAryDestruct
  loc_00DED4C8: ret
  loc_00DED4C9: mov ecx, var_10
  loc_00DED4CC: pop edi
  loc_00DED4CD: pop esi
  loc_00DED4CE: xor eax, eax
  loc_00DED4D0: mov fs:[00000000h], ecx
  loc_00DED4D7: pop ebx
  loc_00DED4D8: mov esp, ebp
  loc_00DED4DA: pop ebp
  loc_00DED4DB: retn 002Ch
End Function

Private Function Proc_2_106_DED4F0(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28) 'DED4F0
  loc_00DED4F0: push ebp
  loc_00DED4F1: mov ebp, esp
  loc_00DED4F3: sub esp, 00000008h
  loc_00DED4F6: push 00402836h ; __vbaExceptHandler
  loc_00DED4FB: mov eax, fs:[00000000h]
  loc_00DED501: push eax
  loc_00DED502: mov fs:[00000000h], esp
  loc_00DED509: sub esp, 00000160h
  loc_00DED50F: push ebx
  loc_00DED510: push esi
  loc_00DED511: push edi
  loc_00DED512: mov var_8, esp
  loc_00DED515: mov var_4, 00401370h
  loc_00DED51C: mov ecx, 0000000Bh
  loc_00DED521: xor eax, eax
  loc_00DED523: lea edi, var_6C
  loc_00DED526: mov var_8C, al
  loc_00DED52C: repz stosd
  loc_00DED52E: mov eax, arg_20
  loc_00DED531: lea ecx, var_24
  loc_00DED534: xor ebx, ebx
  loc_00DED536: push eax
  loc_00DED537: push ecx
  loc_00DED538: mov var_18, bl
  loc_00DED53B: mov var_24, ebx
  loc_00DED53E: mov var_3C, ebx
  loc_00DED541: mov var_40, ebx
  loc_00DED544: mov var_74, ebx
  loc_00DED547: mov var_9C, ebx
  loc_00DED54D: mov var_A4, ebx
  loc_00DED553: mov var_B4, ebx
  loc_00DED559: mov var_C4, ebx
  loc_00DED55F: mov var_D4, ebx
  loc_00DED565: mov var_D8, ebx
  loc_00DED56B: mov var_DC, ebx
  loc_00DED571: mov var_F0, ebx
  loc_00DED577: call [004010B4h] ; __vbaObjSetAddref
  loc_00DED57D: mov esi, arg_18
  loc_00DED580: cmp esi, ebx
  loc_00DED582: jz 00DEEC62h
  loc_00DED588: mov ebx, arg_1C
  loc_00DED58B: test ebx, ebx
  loc_00DED58D: jz 00DEEC62h
  loc_00DED593: mov edx, var_24
  loc_00DED596: push 00000000h
  loc_00DED598: push edx
  loc_00DED599: call [0040114Ch] ; __vbaObjIs
  loc_00DED59F: test ax, ax
  loc_00DED5A2: jnz 00DEEC62h
  loc_00DED5A8: mov edi, Me
  loc_00DED5AB: mov eax, [edi+00000048h]
  loc_00DED5AE: cmp eax, 00000001h
  loc_00DED5B1: jnz 00DED5BEh
  loc_00DED5B3: mov eax, [edi+00000168h]
  loc_00DED5B9: mov var_3C, eax
  loc_00DED5BC: jmp 00DED5CCh
  loc_00DED5BE: cmp eax, 00000002h
  loc_00DED5C1: jnz 00DED5CCh
  loc_00DED5C3: mov ecx, [edi+0000016Ch]
  loc_00DED5C9: mov var_3C, ecx
  loc_00DED5CC: cmp [edi+0000007Ch], 0000h
  loc_00DED5D1: jnz 00DED637h
  loc_00DED5D3: mov eax, [edi+0000015Ch]
  loc_00DED5D9: sub eax, 00000000h
  loc_00DED5DC: jz 00DED626h
  loc_00DED5DE: dec eax
  loc_00DED5DF: jnz 00DED637h
  loc_00DED5E1: xor edx, edx
  loc_00DED5E3: mov dl, [edi+00000158h]
  loc_00DED5E9: mov var_114, edx
  loc_00DED5EF: fild real4 ptr var_114
  loc_00DED5F5: fstp real8 ptr var_11C
  loc_00DED5FB: fld real8 ptr var_11C
  loc_00DED601: fmul st0, real8 ptr [00401358h]
  loc_00DED607: fnstsw ax
  loc_00DED609: test al, 0Dh
  loc_00DED60B: jnz 00DEECC9h
  loc_00DED611: call [0040111Ch] ; __vbaFpUI1
  loc_00DED617: mov var_8C, al
  loc_00DED61D: mov arg_28, FFFFFFFFh
  loc_00DED624: jmp 00DED637h
  loc_00DED626: mov ecx, 00000034h
  loc_00DED62B: call [00401144h] ; __vbaUI1I2
  loc_00DED631: mov var_8C, al
  loc_00DED637: cmp [edi+00000048h], 00000001h
  loc_00DED63B: jnz 00DED646h
  loc_00DED63D: mov al, [edi+00000159h]
  loc_00DED643: mov var_18, al
  loc_00DED646: mov ecx, [edi]
  loc_00DED648: lea edx, var_DC
  loc_00DED64E: push edx
  loc_00DED64F: push edi
  loc_00DED650: call [ecx+000000D8h]
  loc_00DED656: test eax, eax
  loc_00DED658: fnclex
  loc_00DED65A: jge 00DED66Eh
  loc_00DED65C: push 000000D8h
  loc_00DED661: push 006DCB00h
  loc_00DED666: push edi
  loc_00DED667: push eax
  loc_00DED668: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DED66E: mov eax, var_DC
  loc_00DED674: push eax
  loc_00DED675: call 006DD3F0h ; CreateCompatibleDC(%x1v)
  loc_00DED67A: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DED680: mov var_E0, eax
  loc_00DED686: call edi
  loc_00DED688: test esi, esi
  loc_00DED68A: jge 00DED7B4h
  loc_00DED690: mov ecx, var_24
  loc_00DED693: push 00000000h
  loc_00DED695: push 00000004h
  loc_00DED697: lea edx, var_B4
  loc_00DED69D: push ecx
  loc_00DED69E: push edx
  loc_00DED69F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DED6A5: mov esi, Me
  loc_00DED6A8: add esp, 00000010h
  loc_00DED6AB: lea edx, var_D8
  loc_00DED6B1: mov eax, [esi+00000010h]
  loc_00DED6B4: push edx
  loc_00DED6B5: push eax
  loc_00DED6B6: mov ecx, [eax]
  loc_00DED6B8: call [ecx+00000110h]
  loc_00DED6BE: test eax, eax
  loc_00DED6C0: fnclex
  loc_00DED6C2: jge 00DED6D9h
  loc_00DED6C4: mov ecx, [esi+00000010h]
  loc_00DED6C7: push 00000110h
  loc_00DED6CC: push 006DCB00h
  loc_00DED6D1: push ecx
  loc_00DED6D2: push eax
  loc_00DED6D3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DED6D9: mov dx, var_D8
  loc_00DED6E0: mov eax, 00000002h
  loc_00DED6E5: mov var_CC, dx
  loc_00DED6EC: mov edx, [esi+00000010h]
  loc_00DED6EF: mov var_C4, eax
  loc_00DED6F5: mov ecx, 00000008h
  loc_00DED6FA: mov esi, [edx]
  loc_00DED6FC: lea edx, var_DC
  loc_00DED702: push edx
  loc_00DED703: sub esp, 00000010h
  loc_00DED706: mov edx, esp
  loc_00DED708: sub esp, 00000010h
  loc_00DED70B: mov [edx], eax
  loc_00DED70D: mov eax, var_D0
  loc_00DED713: mov [edx+00000004h], eax
  loc_00DED716: mov eax, var_CC
  loc_00DED71C: mov [edx+00000008h], eax
  loc_00DED71F: mov eax, var_C8
  loc_00DED725: mov [edx+0000000Ch], eax
  loc_00DED728: mov eax, var_C4
  loc_00DED72E: mov edx, esp
  loc_00DED730: mov [edx], eax
  loc_00DED732: mov eax, var_C0
  loc_00DED738: mov [edx+00000004h], eax
  loc_00DED73B: mov [edx+00000008h], ecx
  loc_00DED73E: mov ecx, var_B8
  loc_00DED744: mov [edx+0000000Ch], ecx
  loc_00DED747: lea edx, var_B4
  loc_00DED74D: push edx
  loc_00DED74E: call [004011D0h] ; __vbaI4Var
  loc_00DED754: mov var_120, eax
  loc_00DED75A: mov edx, esi
  loc_00DED75C: fild real4 ptr var_120
  loc_00DED762: mov esi, Me
  loc_00DED765: fstp real4 ptr var_124
  loc_00DED76B: mov eax, var_124
  loc_00DED771: mov ecx, [esi+00000010h]
  loc_00DED774: push eax
  loc_00DED775: push ecx
  loc_00DED776: call [edx+00000374h]
  loc_00DED77C: test eax, eax
  loc_00DED77E: fnclex
  loc_00DED780: jge 00DED797h
  loc_00DED782: mov ecx, [esi+00000010h]
  loc_00DED785: push 00000374h
  loc_00DED78A: push 006DCB00h
  loc_00DED78F: push ecx
  loc_00DED790: push eax
  loc_00DED791: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DED797: fld real4 ptr var_DC
  loc_00DED79D: call [00401208h] ; __vbaFpI4
  loc_00DED7A3: mov esi, eax
  loc_00DED7A5: lea ecx, var_B4
  loc_00DED7AB: mov arg_18, esi
  loc_00DED7AE: call [00401024h] ; __vbaFreeVar
  loc_00DED7B4: test ebx, ebx
  loc_00DED7B6: jge 00DED8E0h
  loc_00DED7BC: mov edx, var_24
  loc_00DED7BF: push 00000000h
  loc_00DED7C1: push 00000005h
  loc_00DED7C3: lea eax, var_B4
  loc_00DED7C9: push edx
  loc_00DED7CA: push eax
  loc_00DED7CB: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DED7D1: mov ebx, Me
  loc_00DED7D4: add esp, 00000010h
  loc_00DED7D7: lea edx, var_D8
  loc_00DED7DD: mov eax, [ebx+00000010h]
  loc_00DED7E0: push edx
  loc_00DED7E1: push eax
  loc_00DED7E2: mov ecx, [eax]
  loc_00DED7E4: call [ecx+00000110h]
  loc_00DED7EA: test eax, eax
  loc_00DED7EC: fnclex
  loc_00DED7EE: jge 00DED805h
  loc_00DED7F0: mov ecx, [ebx+00000010h]
  loc_00DED7F3: push 00000110h
  loc_00DED7F8: push 006DCB00h
  loc_00DED7FD: push ecx
  loc_00DED7FE: push eax
  loc_00DED7FF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DED805: mov dx, var_D8
  loc_00DED80C: mov eax, 00000002h
  loc_00DED811: mov var_CC, dx
  loc_00DED818: mov edx, [ebx+00000010h]
  loc_00DED81B: mov var_C4, eax
  loc_00DED821: mov ecx, 00000008h
  loc_00DED826: mov ebx, [edx]
  loc_00DED828: lea edx, var_DC
  loc_00DED82E: push edx
  loc_00DED82F: sub esp, 00000010h
  loc_00DED832: mov edx, esp
  loc_00DED834: sub esp, 00000010h
  loc_00DED837: mov [edx], eax
  loc_00DED839: mov eax, var_D0
  loc_00DED83F: mov [edx+00000004h], eax
  loc_00DED842: mov eax, var_CC
  loc_00DED848: mov [edx+00000008h], eax
  loc_00DED84B: mov eax, var_C8
  loc_00DED851: mov [edx+0000000Ch], eax
  loc_00DED854: mov eax, var_C4
  loc_00DED85A: mov edx, esp
  loc_00DED85C: mov [edx], eax
  loc_00DED85E: mov eax, var_C0
  loc_00DED864: mov [edx+00000004h], eax
  loc_00DED867: mov [edx+00000008h], ecx
  loc_00DED86A: mov ecx, var_B8
  loc_00DED870: mov [edx+0000000Ch], ecx
  loc_00DED873: lea edx, var_B4
  loc_00DED879: push edx
  loc_00DED87A: call [004011D0h] ; __vbaI4Var
  loc_00DED880: mov var_12C, eax
  loc_00DED886: mov edx, ebx
  loc_00DED888: fild real4 ptr var_12C
  loc_00DED88E: mov ebx, Me
  loc_00DED891: fstp real4 ptr var_130
  loc_00DED897: mov eax, var_130
  loc_00DED89D: mov ecx, [ebx+00000010h]
  loc_00DED8A0: push eax
  loc_00DED8A1: push ecx
  loc_00DED8A2: call [edx+00000378h]
  loc_00DED8A8: test eax, eax
  loc_00DED8AA: fnclex
  loc_00DED8AC: jge 00DED8C3h
  loc_00DED8AE: mov ecx, [ebx+00000010h]
  loc_00DED8B1: push 00000378h
  loc_00DED8B6: push 006DCB00h
  loc_00DED8BB: push ecx
  loc_00DED8BC: push eax
  loc_00DED8BD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DED8C3: fld real4 ptr var_DC
  loc_00DED8C9: call [00401208h] ; __vbaFpI4
  loc_00DED8CF: mov ebx, eax
  loc_00DED8D1: lea ecx, var_B4
  loc_00DED8D7: mov arg_1C, ebx
  loc_00DED8DA: call [00401024h] ; __vbaFreeVar
  loc_00DED8E0: mov edx, var_24
  loc_00DED8E3: push 00000000h
  loc_00DED8E5: push 00000000h
  loc_00DED8E7: lea eax, var_B4
  loc_00DED8ED: push edx
  loc_00DED8EE: push eax
  loc_00DED8EF: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DED8F5: add esp, 00000010h
  loc_00DED8F8: push eax
  loc_00DED8F9: call [004011D0h] ; __vbaI4Var
  loc_00DED8FF: mov ecx, var_E0
  loc_00DED905: push eax
  loc_00DED906: push ecx
  loc_00DED907: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DED90C: mov var_DC, eax
  loc_00DED912: call edi
  loc_00DED914: mov edx, var_DC
  loc_00DED91A: lea ecx, var_B4
  loc_00DED920: mov var_98, edx
  loc_00DED926: call [00401024h] ; __vbaFreeVar
  loc_00DED92C: mov eax, var_E0
  loc_00DED932: push eax
  loc_00DED933: call 006DD3F0h ; CreateCompatibleDC(%x1v)
  loc_00DED938: mov var_DC, eax
  loc_00DED93E: call edi
  loc_00DED940: mov edx, var_E0
  loc_00DED946: mov ecx, var_DC
  loc_00DED94C: push edx
  loc_00DED94D: mov var_84, ecx
  loc_00DED953: call 006DD3F0h ; CreateCompatibleDC(%x1v)
  loc_00DED958: mov var_DC, eax
  loc_00DED95E: call edi
  loc_00DED960: mov ecx, arg_C
  loc_00DED963: mov eax, var_DC
  loc_00DED969: push ebx
  loc_00DED96A: push esi
  loc_00DED96B: push ecx
  loc_00DED96C: mov var_38, eax
  loc_00DED96F: call 006DD440h ; CreateCompatibleBitmap(%x1v, %x2v, %x3v)
  loc_00DED974: mov var_DC, eax
  loc_00DED97A: call edi
  loc_00DED97C: mov eax, arg_C
  loc_00DED97F: mov edx, var_DC
  loc_00DED985: push ebx
  loc_00DED986: push esi
  loc_00DED987: push eax
  loc_00DED988: mov var_78, edx
  loc_00DED98B: call 006DD440h ; CreateCompatibleBitmap(%x1v, %x2v, %x3v)
  loc_00DED990: mov var_DC, eax
  loc_00DED996: call edi
  loc_00DED998: mov edx, var_78
  loc_00DED99B: mov eax, var_84
  loc_00DED9A1: mov ecx, var_DC
  loc_00DED9A7: push edx
  loc_00DED9A8: push eax
  loc_00DED9A9: mov var_80, ecx
  loc_00DED9AC: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DED9B1: mov var_DC, eax
  loc_00DED9B7: call edi
  loc_00DED9B9: mov edx, var_80
  loc_00DED9BC: mov eax, var_38
  loc_00DED9BF: mov ecx, var_DC
  loc_00DED9C5: push edx
  loc_00DED9C6: push eax
  loc_00DED9C7: mov var_20, ecx
  loc_00DED9CA: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DED9CF: mov var_DC, eax
  loc_00DED9D5: call edi
  loc_00DED9D7: mov eax, esi
  loc_00DED9D9: mov ecx, var_DC
  loc_00DED9DF: imul eax, eax, 00000020h
  loc_00DED9E2: jo 00DEECCEh
  loc_00DED9E8: add eax, 0000001Fh
  loc_00DED9EB: mov var_30, ecx
  loc_00DED9EE: jo 00DEECCEh
  loc_00DED9F4: mov var_6C, 00000028h
  loc_00DED9FB: mov var_68, esi
  loc_00DED9FE: mov var_64, ebx
  loc_00DEDA01: mov var_60, 0001h
  loc_00DEDA07: mov var_5E, 0020h
  loc_00DEDA0D: cdq
  loc_00DEDA0E: and edx, 0000001Fh
  loc_00DEDA11: push 00000000h
  loc_00DEDA13: add eax, edx
  loc_00DEDA15: lea edx, var_40
  loc_00DEDA18: sar eax, 05h
  loc_00DEDA1B: imul eax, eax, 00000004h
  loc_00DEDA1E: jo 00DEECCEh
  loc_00DEDA24: imul eax, ebx
  loc_00DEDA27: jo 00DEECCEh
  loc_00DEDA2D: mov var_58, eax
  loc_00DEDA30: sub eax, 00000001h
  loc_00DEDA33: jo 00DEECCEh
  loc_00DEDA39: push eax
  loc_00DEDA3A: push 00000001h
  loc_00DEDA3C: push 00000000h
  loc_00DEDA3E: push edx
  loc_00DEDA3F: push 00000004h
  loc_00DEDA41: push 00000000h
  loc_00DEDA43: call [00401138h] ; __vbaRedim
  loc_00DEDA49: mov eax, var_40
  loc_00DEDA4C: add esp, 0000001Ch
  loc_00DEDA4F: push 00000000h
  loc_00DEDA51: push eax
  loc_00DEDA52: push 00000001h
  loc_00DEDA54: call [0040117Ch] ; __vbaUbound
  loc_00DEDA5A: push eax
  loc_00DEDA5B: push 00000001h
  loc_00DEDA5D: lea ecx, var_9C
  loc_00DEDA63: push 00000000h
  loc_00DEDA65: push ecx
  loc_00DEDA66: push 00000004h
  loc_00DEDA68: push 00000000h
  loc_00DEDA6A: call [00401138h] ; __vbaRedim
  loc_00DEDA70: mov edx, arg_14
  loc_00DEDA73: mov eax, arg_10
  loc_00DEDA76: mov ecx, arg_C
  loc_00DEDA79: add esp, 0000001Ch
  loc_00DEDA7C: push 00CC0020h
  loc_00DEDA81: push edx
  loc_00DEDA82: mov edx, var_84
  loc_00DEDA88: push eax
  loc_00DEDA89: push ecx
  loc_00DEDA8A: push ebx
  loc_00DEDA8B: push esi
  loc_00DEDA8C: push 00000000h
  loc_00DEDA8E: push 00000000h
  loc_00DEDA90: push edx
  loc_00DEDA91: call 006DD820h ; BitBlt(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v)
  loc_00DEDA96: call edi
  loc_00DEDA98: mov eax, var_E0
  loc_00DEDA9E: mov ecx, var_38
  loc_00DEDAA1: push 00CC0020h
  loc_00DEDAA6: push 00000000h
  loc_00DEDAA8: push 00000000h
  loc_00DEDAAA: push eax
  loc_00DEDAAB: push ebx
  loc_00DEDAAC: push esi
  loc_00DEDAAD: push 00000000h
  loc_00DEDAAF: push 00000000h
  loc_00DEDAB1: push ecx
  loc_00DEDAB2: call 006DD820h ; BitBlt(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v)
  loc_00DEDAB7: call edi
  loc_00DEDAB9: mov edx, var_40
  loc_00DEDABC: lea eax, var_A4
  loc_00DEDAC2: push edx
  loc_00DEDAC3: push eax
  loc_00DEDAC4: call [004011DCh] ; __vbaAryLock
  loc_00DEDACA: mov ecx, var_A4
  loc_00DEDAD0: test ecx, ecx
  loc_00DEDAD2: jz 00DEDB03h
  loc_00DEDAD4: cmp [ecx], 0001h
  loc_00DEDAD8: jnz 00DEDB03h
  loc_00DEDADA: mov eax, [ecx+00000014h]
  loc_00DEDADD: mov edx, [ecx+00000010h]
  loc_00DEDAE0: neg eax
  loc_00DEDAE2: cmp eax, edx
  loc_00DEDAE4: mov var_E4, eax
  loc_00DEDAEA: jb 00DEDAFEh
  loc_00DEDAEC: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDAF2: mov ecx, var_A4
  loc_00DEDAF8: mov eax, var_E4
  loc_00DEDAFE: shl eax, 02h
  loc_00DEDB01: jmp 00DEDB0Fh
  loc_00DEDB03: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDB09: mov ecx, var_A4
  loc_00DEDB0F: mov ecx, [ecx+0000000Ch]
  loc_00DEDB12: lea edx, var_6C
  loc_00DEDB15: push 00000000h
  loc_00DEDB17: add ecx, eax
  loc_00DEDB19: mov eax, var_84
  loc_00DEDB1F: push edx
  loc_00DEDB20: mov edx, var_78
  loc_00DEDB23: push ecx
  loc_00DEDB24: push ebx
  loc_00DEDB25: push 00000000h
  loc_00DEDB27: push edx
  loc_00DEDB28: push eax
  loc_00DEDB29: call 006DD2F4h ; GetDIBits(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v)
  loc_00DEDB2E: call edi
  loc_00DEDB30: lea ecx, var_A4
  loc_00DEDB36: push ecx
  loc_00DEDB37: call [00401244h] ; __vbaAryUnlock
  loc_00DEDB3D: mov edx, var_9C
  loc_00DEDB43: lea eax, var_A4
  loc_00DEDB49: push edx
  loc_00DEDB4A: push eax
  loc_00DEDB4B: call [004011DCh] ; __vbaAryLock
  loc_00DEDB51: mov ecx, var_A4
  loc_00DEDB57: test ecx, ecx
  loc_00DEDB59: jz 00DEDB8Ah
  loc_00DEDB5B: cmp [ecx], 0001h
  loc_00DEDB5F: jnz 00DEDB8Ah
  loc_00DEDB61: mov eax, [ecx+00000014h]
  loc_00DEDB64: mov edx, [ecx+00000010h]
  loc_00DEDB67: neg eax
  loc_00DEDB69: cmp eax, edx
  loc_00DEDB6B: mov var_E4, eax
  loc_00DEDB71: jb 00DEDB85h
  loc_00DEDB73: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDB79: mov ecx, var_A4
  loc_00DEDB7F: mov eax, var_E4
  loc_00DEDB85: shl eax, 02h
  loc_00DEDB88: jmp 00DEDB96h
  loc_00DEDB8A: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDB90: mov ecx, var_A4
  loc_00DEDB96: mov ecx, [ecx+0000000Ch]
  loc_00DEDB99: lea edx, var_6C
  loc_00DEDB9C: push 00000000h
  loc_00DEDB9E: add ecx, eax
  loc_00DEDBA0: mov eax, var_38
  loc_00DEDBA3: push edx
  loc_00DEDBA4: mov edx, var_80
  loc_00DEDBA7: push ecx
  loc_00DEDBA8: push ebx
  loc_00DEDBA9: push 00000000h
  loc_00DEDBAB: push edx
  loc_00DEDBAC: push eax
  loc_00DEDBAD: call 006DD2F4h ; GetDIBits(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v)
  loc_00DEDBB2: call edi
  loc_00DEDBB4: lea ecx, var_A4
  loc_00DEDBBA: push ecx
  loc_00DEDBBB: call [00401244h] ; __vbaAryUnlock
  loc_00DEDBC1: mov eax, arg_24
  loc_00DEDBC4: cmp eax, FFFFFFFFh
  loc_00DEDBC7: jz 00DEDC2Ch
  loc_00DEDBC9: cdq
  loc_00DEDBCA: and edx, 0000FFFFh
  loc_00DEDBD0: add eax, edx
  loc_00DEDBD2: mov ecx, eax
  loc_00DEDBD4: sar ecx, 10h
  loc_00DEDBD7: and ecx, 800000FFh
  loc_00DEDBDD: jns 00DEDBE7h
  loc_00DEDBDF: dec ecx
  loc_00DEDBE0: or ecx, FFFFFF00h
  loc_00DEDBE6: inc ecx
  loc_00DEDBE7: call [00401158h] ; __vbaUI1I4
  loc_00DEDBED: mov var_74, al
  loc_00DEDBF0: mov eax, arg_24
  loc_00DEDBF3: cdq
  loc_00DEDBF4: and edx, 000000FFh
  loc_00DEDBFA: add eax, edx
  loc_00DEDBFC: mov ecx, eax
  loc_00DEDBFE: sar ecx, 08h
  loc_00DEDC01: and ecx, 800000FFh
  loc_00DEDC07: jns 00DEDC11h
  loc_00DEDC09: dec ecx
  loc_00DEDC0A: or ecx, FFFFFF00h
  loc_00DEDC10: inc ecx
  loc_00DEDC11: call [00401158h] ; __vbaUI1I4
  loc_00DEDC17: mov ecx, arg_24
  loc_00DEDC1A: mov var_73, al
  loc_00DEDC1D: and ecx, 000000FFh
  loc_00DEDC23: call [00401158h] ; __vbaUI1I4
  loc_00DEDC29: mov var_72, al
  loc_00DEDC2C: mov ecx, esi
  loc_00DEDC2E: mov eax, ebx
  loc_00DEDC30: sub ecx, 00000001h
  loc_00DEDC33: jo 00DEECCEh
  loc_00DEDC39: sub eax, 00000001h
  loc_00DEDC3C: mov var_34, ecx
  loc_00DEDC3F: jo 00DEECCEh
  loc_00DEDC45: xor edx, edx
  loc_00DEDC47: mov var_104, eax
  loc_00DEDC4D: mov var_A0, edx
  loc_00DEDC53: cmp edx, eax
  loc_00DEDC55: jg 00DEEADAh
  loc_00DEDC5B: mov eax, edx
  loc_00DEDC5D: mov var_70, 00000000h
  loc_00DEDC64: imul eax, esi
  loc_00DEDC67: jo 00DEECCEh
  loc_00DEDC6D: mov var_94, eax
  loc_00DEDC73: cmp var_70, ecx
  loc_00DEDC76: jg 00DEEABAh
  loc_00DEDC7C: mov edx, var_70
  loc_00DEDC7F: mov esi, eax
  loc_00DEDC81: mov eax, Me
  loc_00DEDC84: add esi, edx
  loc_00DEDC86: jo 00DEECCEh
  loc_00DEDC8C: cmp [eax+0000007Ch], 0000h
  loc_00DEDC91: mov var_14, esi
  loc_00DEDC94: jz 00DEDD6Ch
  loc_00DEDC9A: mov ecx, [eax+00000048h]
  loc_00DEDC9D: cmp ecx, 00000001h
  loc_00DEDCA0: mov ecx, var_9C
  loc_00DEDCA6: jnz 00DEDD01h
  loc_00DEDCA8: test ecx, ecx
  loc_00DEDCAA: jz 00DEDCE4h
  loc_00DEDCAC: cmp [ecx], 0001h
  loc_00DEDCB0: jnz 00DEDCE4h
  loc_00DEDCB2: mov edx, [ecx+00000014h]
  loc_00DEDCB5: mov eax, [ecx+00000010h]
  loc_00DEDCB8: mov edi, esi
  loc_00DEDCBA: sub edi, edx
  loc_00DEDCBC: cmp edi, eax
  loc_00DEDCBE: jb 00DEDCCCh
  loc_00DEDCC0: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDCC6: mov ecx, var_9C
  loc_00DEDCCC: mov edx, [ecx+0000000Ch]
  loc_00DEDCCF: lea eax, [edi*4]
  loc_00DEDCD6: xor ecx, ecx
  loc_00DEDCD8: mov cl, [edx+eax+00000003h]
  loc_00DEDCDC: mov eax, var_18
  loc_00DEDCDF: jmp 00DEDDBCh
  loc_00DEDCE4: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDCEA: mov ecx, var_9C
  loc_00DEDCF0: mov edx, [ecx+0000000Ch]
  loc_00DEDCF3: xor ecx, ecx
  loc_00DEDCF5: mov cl, [edx+eax+00000003h]
  loc_00DEDCF9: mov eax, var_18
  loc_00DEDCFC: jmp 00DEDDBCh
  loc_00DEDD01: test ecx, ecx
  loc_00DEDD03: jz 00DEDD2Eh
  loc_00DEDD05: cmp [ecx], 0001h
  loc_00DEDD09: jnz 00DEDD2Eh
  loc_00DEDD0B: mov edx, [ecx+00000014h]
  loc_00DEDD0E: mov eax, [ecx+00000010h]
  loc_00DEDD11: mov edi, esi
  loc_00DEDD13: sub edi, edx
  loc_00DEDD15: cmp edi, eax
  loc_00DEDD17: jb 00DEDD25h
  loc_00DEDD19: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDD1F: mov ecx, var_9C
  loc_00DEDD25: lea eax, [edi*4]
  loc_00DEDD2C: jmp 00DEDD3Ah
  loc_00DEDD2E: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDD34: mov ecx, var_9C
  loc_00DEDD3A: mov edx, [ecx+0000000Ch]
  loc_00DEDD3D: xor ecx, ecx
  loc_00DEDD3F: mov cl, [edx+eax+00000003h]
  loc_00DEDD43: mov edx, Me
  loc_00DEDD46: xor eax, eax
  loc_00DEDD48: mov al, [edx+00000158h]
  loc_00DEDD4E: imul ecx, eax
  loc_00DEDD51: mov eax, 80808081h
  loc_00DEDD56: jo 00DEECCEh
  loc_00DEDD5C: imul ecx
  loc_00DEDD5E: add edx, ecx
  loc_00DEDD60: sar edx, 07h
  loc_00DEDD63: mov eax, edx
  loc_00DEDD65: shr eax, 1Fh
  loc_00DEDD68: add edx, eax
  loc_00DEDD6A: jmp 00DEDDDDh
  loc_00DEDD6C: mov ecx, var_9C
  loc_00DEDD72: test ecx, ecx
  loc_00DEDD74: jz 00DEDD9Fh
  loc_00DEDD76: cmp [ecx], 0001h
  loc_00DEDD7A: jnz 00DEDD9Fh
  loc_00DEDD7C: mov edx, [ecx+00000014h]
  loc_00DEDD7F: mov eax, [ecx+00000010h]
  loc_00DEDD82: mov edi, esi
  loc_00DEDD84: sub edi, edx
  loc_00DEDD86: cmp edi, eax
  loc_00DEDD88: jb 00DEDD96h
  loc_00DEDD8A: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDD90: mov ecx, var_9C
  loc_00DEDD96: lea eax, [edi*4]
  loc_00DEDD9D: jmp 00DEDDABh
  loc_00DEDD9F: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDDA5: mov ecx, var_9C
  loc_00DEDDAB: mov ecx, [ecx+0000000Ch]
  loc_00DEDDAE: xor edx, edx
  loc_00DEDDB0: mov dl, [ecx+eax+00000003h]
  loc_00DEDDB4: mov eax, var_8C
  loc_00DEDDBA: mov ecx, edx
  loc_00DEDDBC: and eax, 000000FFh
  loc_00DEDDC1: imul ecx, eax
  loc_00DEDDC4: mov eax, 80808081h
  loc_00DEDDC9: jo 00DEECCEh
  loc_00DEDDCF: imul ecx
  loc_00DEDDD1: add edx, ecx
  loc_00DEDDD3: sar edx, 07h
  loc_00DEDDD6: mov ecx, edx
  loc_00DEDDD8: shr ecx, 1Fh
  loc_00DEDDDB: add edx, ecx
  loc_00DEDDDD: mov edi, edx
  loc_00DEDDDF: mov edx, var_40
  loc_00DEDDE2: mov ebx, 000000FFh
  loc_00DEDDE7: lea eax, var_F0
  loc_00DEDDED: sub ebx, edi
  loc_00DEDDEF: push edx
  loc_00DEDDF0: jo 00DEECCEh
  loc_00DEDDF6: push eax
  loc_00DEDDF7: mov var_88, ebx
  loc_00DEDDFD: call [004011DCh] ; __vbaAryLock
  loc_00DEDE03: mov ecx, var_F0
  loc_00DEDE09: test ecx, ecx
  loc_00DEDE0B: jz 00DEDE34h
  loc_00DEDE0D: cmp [ecx], 0001h
  loc_00DEDE11: jnz 00DEDE34h
  loc_00DEDE13: mov edx, [ecx+00000014h]
  loc_00DEDE16: mov eax, [ecx+00000010h]
  loc_00DEDE19: sub esi, edx
  loc_00DEDE1B: cmp esi, eax
  loc_00DEDE1D: jb 00DEDE2Bh
  loc_00DEDE1F: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDE25: mov ecx, var_F0
  loc_00DEDE2B: lea eax, [esi*4]
  loc_00DEDE32: jmp 00DEDE40h
  loc_00DEDE34: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDE3A: mov ecx, var_F0
  loc_00DEDE40: mov esi, [ecx+0000000Ch]
  loc_00DEDE43: add esi, eax
  loc_00DEDE45: mov eax, arg_24
  loc_00DEDE48: cmp eax, FFFFFFFFh
  loc_00DEDE4B: jz 00DEDF4Ch
  loc_00DEDE51: cmp edi, 000000FFh
  loc_00DEDE57: jnz 00DEDEA9h
  loc_00DEDE59: mov ecx, var_40
  loc_00DEDE5C: test ecx, ecx
  loc_00DEDE5E: jz 00DEDE89h
  loc_00DEDE60: cmp [ecx], 0001h
  loc_00DEDE64: jnz 00DEDE89h
  loc_00DEDE66: mov eax, var_14
  loc_00DEDE69: mov edx, [ecx+00000014h]
  loc_00DEDE6C: sub eax, edx
  loc_00DEDE6E: mov esi, eax
  loc_00DEDE70: mov eax, [ecx+00000010h]
  loc_00DEDE73: cmp esi, eax
  loc_00DEDE75: jb 00DEDE80h
  loc_00DEDE77: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDE7D: mov ecx, var_40
  loc_00DEDE80: lea eax, [esi*4]
  loc_00DEDE87: jmp 00DEDE92h
  loc_00DEDE89: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDE8F: mov ecx, var_40
  loc_00DEDE92: mov ecx, [ecx+0000000Ch]
  loc_00DEDE95: lea edx, var_74
  loc_00DEDE98: add ecx, eax
  loc_00DEDE9A: push edx
  loc_00DEDE9B: push ecx
  loc_00DEDE9C: push 00000004h
  loc_00DEDE9E: call [00401060h] ; __vbaCopyBytes
  loc_00DEDEA4: jmp 00DEEA7Ah
  loc_00DEDEA9: test edi, edi
  loc_00DEDEAB: jle 00DEEA7Ah
  loc_00DEDEB1: mov eax, var_72
  loc_00DEDEB4: and eax, 000000FFh
  loc_00DEDEB9: imul eax, edi
  loc_00DEDEBC: jo 00DEECCEh
  loc_00DEDEC2: xor edx, edx
  loc_00DEDEC4: mov dl, [esi+00000002h]
  loc_00DEDEC7: imul edx, ebx
  loc_00DEDECA: jo 00DEECCEh
  loc_00DEDED0: add eax, edx
  loc_00DEDED2: jo 00DEECCEh
  loc_00DEDED8: cdq
  loc_00DEDED9: and edx, 000000FFh
  loc_00DEDEDF: add eax, edx
  loc_00DEDEE1: mov ecx, eax
  loc_00DEDEE3: sar ecx, 08h
  loc_00DEDEE6: call [00401158h] ; __vbaUI1I4
  loc_00DEDEEC: mov [esi+00000002h], al
  loc_00DEDEEF: xor eax, eax
  loc_00DEDEF1: mov al, [esi+00000001h]
  loc_00DEDEF4: mov ecx, var_73
  loc_00DEDEF7: imul eax, ebx
  loc_00DEDEFA: jo 00DEECCEh
  loc_00DEDF00: and ecx, 000000FFh
  loc_00DEDF06: imul ecx, edi
  loc_00DEDF09: jo 00DEECCEh
  loc_00DEDF0F: add eax, ecx
  loc_00DEDF11: jo 00DEECCEh
  loc_00DEDF17: cdq
  loc_00DEDF18: and edx, 000000FFh
  loc_00DEDF1E: add eax, edx
  loc_00DEDF20: mov ecx, eax
  loc_00DEDF22: sar ecx, 08h
  loc_00DEDF25: call [00401158h] ; __vbaUI1I4
  loc_00DEDF2B: mov [esi+00000001h], al
  loc_00DEDF2E: xor eax, eax
  loc_00DEDF30: mov al, [esi]
  loc_00DEDF32: mov edx, var_74
  loc_00DEDF35: imul eax, ebx
  loc_00DEDF38: jo 00DEECCEh
  loc_00DEDF3E: and edx, 000000FFh
  loc_00DEDF44: imul edx, edi
  loc_00DEDF47: jmp 00DEEA56h
  loc_00DEDF4C: cmp arg_28, 0000h
  loc_00DEDF51: jz 00DEE1B3h
  loc_00DEDF57: mov eax, var_9C
  loc_00DEDF5D: test eax, eax
  loc_00DEDF5F: jz 00DEDF91h
  loc_00DEDF61: cmp [eax], 0001h
  loc_00DEDF65: jnz 00DEDF91h
  loc_00DEDF67: mov ebx, var_14
  loc_00DEDF6A: mov edx, [eax+00000014h]
  loc_00DEDF6D: mov ecx, [eax+00000010h]
  loc_00DEDF70: sub ebx, edx
  loc_00DEDF72: cmp ebx, ecx
  loc_00DEDF74: jb 00DEDF82h
  loc_00DEDF76: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDF7C: mov eax, var_9C
  loc_00DEDF82: lea ecx, [ebx*4]
  loc_00DEDF89: mov var_138, ecx
  loc_00DEDF8F: jmp 00DEDFA3h
  loc_00DEDF91: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDF97: mov var_138, eax
  loc_00DEDF9D: mov eax, var_9C
  loc_00DEDFA3: test eax, eax
  loc_00DEDFA5: jz 00DEDFD7h
  loc_00DEDFA7: cmp [eax], 0001h
  loc_00DEDFAB: jnz 00DEDFD7h
  loc_00DEDFAD: mov ebx, var_14
  loc_00DEDFB0: mov edx, [eax+00000014h]
  loc_00DEDFB3: mov ecx, [eax+00000010h]
  loc_00DEDFB6: sub ebx, edx
  loc_00DEDFB8: cmp ebx, ecx
  loc_00DEDFBA: jb 00DEDFC8h
  loc_00DEDFBC: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDFC2: mov eax, var_9C
  loc_00DEDFC8: lea edx, [ebx*4]
  loc_00DEDFCF: mov var_13C, edx
  loc_00DEDFD5: jmp 00DEDFE9h
  loc_00DEDFD7: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEDFDD: mov var_13C, eax
  loc_00DEDFE3: mov eax, var_9C
  loc_00DEDFE9: test eax, eax
  loc_00DEDFEB: jz 00DEE01Dh
  loc_00DEDFED: cmp [eax], 0001h
  loc_00DEDFF1: jnz 00DEE01Dh
  loc_00DEDFF3: mov ebx, var_14
  loc_00DEDFF6: mov edx, [eax+00000014h]
  loc_00DEDFF9: mov ecx, [eax+00000010h]
  loc_00DEDFFC: sub ebx, edx
  loc_00DEDFFE: cmp ebx, ecx
  loc_00DEE000: jb 00DEE00Eh
  loc_00DEE002: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE008: mov eax, var_9C
  loc_00DEE00E: lea ecx, [ebx*4]
  loc_00DEE015: mov var_140, ecx
  loc_00DEE01B: jmp 00DEE02Fh
  loc_00DEE01D: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE023: mov var_140, eax
  loc_00DEE029: mov eax, var_9C
  loc_00DEE02F: mov ebx, [eax+0000000Ch]
  loc_00DEE032: mov ecx, var_138
  loc_00DEE038: mov edx, ebx
  loc_00DEE03A: xor eax, eax
  loc_00DEE03C: mov al, [edx+ecx+00000002h]
  loc_00DEE040: mov var_144, eax
  loc_00DEE046: fild real4 ptr var_144
  loc_00DEE04C: fstp real8 ptr var_14C
  loc_00DEE052: fld real8 ptr var_14C
  loc_00DEE058: fmul st0, real8 ptr [00401350h]
  loc_00DEE05E: fnstsw ax
  loc_00DEE060: test al, 0Dh
  loc_00DEE062: jnz 00DEECC9h
  loc_00DEE068: call [00401208h] ; __vbaFpI4
  loc_00DEE06E: mov var_150, eax
  loc_00DEE074: mov eax, var_13C
  loc_00DEE07A: fild real4 ptr var_150
  loc_00DEE080: xor edx, edx
  loc_00DEE082: xor ecx, ecx
  loc_00DEE084: mov dl, [eax+ebx+00000001h]
  loc_00DEE088: fstp real8 ptr var_158
  loc_00DEE08E: mov var_15C, edx
  loc_00DEE094: mov edx, var_140
  loc_00DEE09A: fild real4 ptr var_15C
  loc_00DEE0A0: mov cl, [edx+ebx]
  loc_00DEE0A3: mov var_168, ecx
  loc_00DEE0A9: fstp real8 ptr var_164
  loc_00DEE0AF: fld real8 ptr var_164
  loc_00DEE0B5: fmul st0, real8 ptr [00401348h]
  loc_00DEE0BB: fadd st0, real8 ptr var_158
  loc_00DEE0C1: fild real4 ptr var_168
  loc_00DEE0C7: fstp real8 ptr var_170
  loc_00DEE0CD: fld real8 ptr var_170
  loc_00DEE0D3: fmul st0, real8 ptr [00401340h]
  loc_00DEE0D9: faddp st1
  loc_00DEE0DB: fnstsw ax
  loc_00DEE0DD: test al, 0Dh
  loc_00DEE0DF: jnz 00DEECC9h
  loc_00DEE0E5: call [00401208h] ; __vbaFpI4
  loc_00DEE0EB: cmp edi, 000000FFh
  loc_00DEE0F1: mov ebx, eax
  loc_00DEE0F3: jnz 00DEE112h
  loc_00DEE0F5: mov edi, [00401158h] ; __vbaUI1I4
  loc_00DEE0FB: mov ecx, ebx
  loc_00DEE0FD: call edi
  loc_00DEE0FF: mov ecx, ebx
  loc_00DEE101: mov [esi+00000002h], al
  loc_00DEE104: call edi
  loc_00DEE106: mov ecx, ebx
  loc_00DEE108: mov [esi+00000001h], al
  loc_00DEE10B: call edi
  loc_00DEE10D: jmp 00DEEA78h
  loc_00DEE112: test edi, edi
  loc_00DEE114: jle 00DEEA7Ah
  loc_00DEE11A: xor eax, eax
  loc_00DEE11C: mov ecx, edi
  loc_00DEE11E: mov al, [esi+00000002h]
  loc_00DEE121: imul eax, var_88
  loc_00DEE128: jo 00DEECCEh
  loc_00DEE12E: imul ecx, ebx
  loc_00DEE131: jo 00DEECCEh
  loc_00DEE137: add eax, ecx
  loc_00DEE139: jo 00DEECCEh
  loc_00DEE13F: cdq
  loc_00DEE140: and edx, 000000FFh
  loc_00DEE146: add eax, edx
  loc_00DEE148: mov ecx, eax
  loc_00DEE14A: sar ecx, 08h
  loc_00DEE14D: call [00401158h] ; __vbaUI1I4
  loc_00DEE153: mov [esi+00000002h], al
  loc_00DEE156: xor eax, eax
  loc_00DEE158: mov al, [esi+00000001h]
  loc_00DEE15B: mov edx, edi
  loc_00DEE15D: imul eax, var_88
  loc_00DEE164: jo 00DEECCEh
  loc_00DEE16A: imul edx, ebx
  loc_00DEE16D: jo 00DEECCEh
  loc_00DEE173: add eax, edx
  loc_00DEE175: jo 00DEECCEh
  loc_00DEE17B: cdq
  loc_00DEE17C: and edx, 000000FFh
  loc_00DEE182: add eax, edx
  loc_00DEE184: mov ecx, eax
  loc_00DEE186: sar ecx, 08h
  loc_00DEE189: call [00401158h] ; __vbaUI1I4
  loc_00DEE18F: mov [esi+00000001h], al
  loc_00DEE192: xor eax, eax
  loc_00DEE194: mov al, [esi]
  loc_00DEE196: imul eax, var_88
  loc_00DEE19D: jo 00DEECCEh
  loc_00DEE1A3: imul edi, ebx
  loc_00DEE1A6: jo 00DEECCEh
  loc_00DEE1AC: add eax, edi
  loc_00DEE1AE: jmp 00DEEA5Eh
  loc_00DEE1B3: cmp edi, 000000FFh
  loc_00DEE1B9: jnz 00DEE4D4h
  loc_00DEE1BF: mov eax, var_3C
  loc_00DEE1C2: cmp eax, 00000001h
  loc_00DEE1C5: jnz 00DEE305h
  loc_00DEE1CB: mov ecx, var_9C
  loc_00DEE1D1: test ecx, ecx
  loc_00DEE1D3: jz 00DEE1FEh
  loc_00DEE1D5: cmp [ecx], ax
  loc_00DEE1D8: jnz 00DEE1FEh
  loc_00DEE1DA: mov edi, var_14
  loc_00DEE1DD: mov edx, [ecx+00000014h]
  loc_00DEE1E0: mov eax, [ecx+00000010h]
  loc_00DEE1E3: sub edi, edx
  loc_00DEE1E5: cmp edi, eax
  loc_00DEE1E7: jb 00DEE1F5h
  loc_00DEE1E9: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE1EF: mov ecx, var_9C
  loc_00DEE1F5: lea eax, [edi*4]
  loc_00DEE1FC: jmp 00DEE20Ah
  loc_00DEE1FE: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE204: mov ecx, var_9C
  loc_00DEE20A: mov ecx, [ecx+0000000Ch]
  loc_00DEE20D: xor edx, edx
  loc_00DEE20F: mov dl, [ecx+eax+00000002h]
  loc_00DEE213: mov edi, edx
  loc_00DEE215: cmp edi, 00000100h
  loc_00DEE21B: jb 00DEE223h
  loc_00DEE21D: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE223: mov eax, Me
  loc_00DEE226: mov ecx, [eax+0000018Ch]
  loc_00DEE22C: mov dl, [ecx+edi]
  loc_00DEE22F: mov [esi+00000002h], dl
  loc_00DEE232: mov ecx, var_9C
  loc_00DEE238: test ecx, ecx
  loc_00DEE23A: jz 00DEE266h
  loc_00DEE23C: cmp [ecx], 0001h
  loc_00DEE240: jnz 00DEE266h
  loc_00DEE242: mov edi, var_14
  loc_00DEE245: mov edx, [ecx+00000014h]
  loc_00DEE248: mov eax, [ecx+00000010h]
  loc_00DEE24B: sub edi, edx
  loc_00DEE24D: cmp edi, eax
  loc_00DEE24F: jb 00DEE25Dh
  loc_00DEE251: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE257: mov ecx, var_9C
  loc_00DEE25D: lea eax, [edi*4]
  loc_00DEE264: jmp 00DEE272h
  loc_00DEE266: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE26C: mov ecx, var_9C
  loc_00DEE272: mov ecx, [ecx+0000000Ch]
  loc_00DEE275: xor edx, edx
  loc_00DEE277: mov dl, [ecx+eax+00000001h]
  loc_00DEE27B: mov edi, edx
  loc_00DEE27D: cmp edi, 00000100h
  loc_00DEE283: jb 00DEE28Bh
  loc_00DEE285: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE28B: mov eax, Me
  loc_00DEE28E: mov ecx, [eax+0000018Ch]
  loc_00DEE294: mov dl, [ecx+edi]
  loc_00DEE297: mov [esi+00000001h], dl
  loc_00DEE29A: mov ecx, var_9C
  loc_00DEE2A0: test ecx, ecx
  loc_00DEE2A2: jz 00DEE2CEh
  loc_00DEE2A4: cmp [ecx], 0001h
  loc_00DEE2A8: jnz 00DEE2CEh
  loc_00DEE2AA: mov edi, var_14
  loc_00DEE2AD: mov edx, [ecx+00000014h]
  loc_00DEE2B0: mov eax, [ecx+00000010h]
  loc_00DEE2B3: sub edi, edx
  loc_00DEE2B5: cmp edi, eax
  loc_00DEE2B7: jb 00DEE2C5h
  loc_00DEE2B9: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE2BF: mov ecx, var_9C
  loc_00DEE2C5: lea eax, [edi*4]
  loc_00DEE2CC: jmp 00DEE2DAh
  loc_00DEE2CE: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE2D4: mov ecx, var_9C
  loc_00DEE2DA: mov ecx, [ecx+0000000Ch]
  loc_00DEE2DD: xor edx, edx
  loc_00DEE2DF: mov dl, [ecx+eax]
  loc_00DEE2E2: mov edi, edx
  loc_00DEE2E4: cmp edi, 00000100h
  loc_00DEE2EA: jb 00DEE2F2h
  loc_00DEE2EC: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE2F2: mov eax, Me
  loc_00DEE2F5: mov ecx, [eax+0000018Ch]
  loc_00DEE2FB: mov dl, [ecx+edi]
  loc_00DEE2FE: mov [esi], dl
  loc_00DEE300: jmp 00DEEA7Ah
  loc_00DEE305: mov ecx, var_9C
  loc_00DEE30B: cmp eax, 00000002h
  loc_00DEE30E: jnz 00DEE449h
  loc_00DEE314: test ecx, ecx
  loc_00DEE316: jz 00DEE342h
  loc_00DEE318: cmp [ecx], 0001h
  loc_00DEE31C: jnz 00DEE342h
  loc_00DEE31E: mov edi, var_14
  loc_00DEE321: mov edx, [ecx+00000014h]
  loc_00DEE324: mov eax, [ecx+00000010h]
  loc_00DEE327: sub edi, edx
  loc_00DEE329: cmp edi, eax
  loc_00DEE32B: jb 00DEE339h
  loc_00DEE32D: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE333: mov ecx, var_9C
  loc_00DEE339: lea eax, [edi*4]
  loc_00DEE340: jmp 00DEE34Eh
  loc_00DEE342: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE348: mov ecx, var_9C
  loc_00DEE34E: mov ecx, [ecx+0000000Ch]
  loc_00DEE351: xor edx, edx
  loc_00DEE353: mov dl, [ecx+eax+00000002h]
  loc_00DEE357: mov edi, edx
  loc_00DEE359: cmp edi, 00000100h
  loc_00DEE35F: jb 00DEE367h
  loc_00DEE361: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE367: mov eax, Me
  loc_00DEE36A: mov ecx, [eax+000001A8h]
  loc_00DEE370: mov dl, [ecx+edi]
  loc_00DEE373: mov [esi+00000002h], dl
  loc_00DEE376: mov ecx, var_9C
  loc_00DEE37C: test ecx, ecx
  loc_00DEE37E: jz 00DEE3AAh
  loc_00DEE380: cmp [ecx], 0001h
  loc_00DEE384: jnz 00DEE3AAh
  loc_00DEE386: mov edi, var_14
  loc_00DEE389: mov edx, [ecx+00000014h]
  loc_00DEE38C: mov eax, [ecx+00000010h]
  loc_00DEE38F: sub edi, edx
  loc_00DEE391: cmp edi, eax
  loc_00DEE393: jb 00DEE3A1h
  loc_00DEE395: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE39B: mov ecx, var_9C
  loc_00DEE3A1: lea eax, [edi*4]
  loc_00DEE3A8: jmp 00DEE3B6h
  loc_00DEE3AA: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE3B0: mov ecx, var_9C
  loc_00DEE3B6: mov ecx, [ecx+0000000Ch]
  loc_00DEE3B9: xor edx, edx
  loc_00DEE3BB: mov dl, [ecx+eax+00000001h]
  loc_00DEE3BF: mov edi, edx
  loc_00DEE3C1: cmp edi, 00000100h
  loc_00DEE3C7: jb 00DEE3CFh
  loc_00DEE3C9: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE3CF: mov eax, Me
  loc_00DEE3D2: mov ecx, [eax+000001A8h]
  loc_00DEE3D8: mov dl, [ecx+edi]
  loc_00DEE3DB: mov [esi+00000001h], dl
  loc_00DEE3DE: mov ecx, var_9C
  loc_00DEE3E4: test ecx, ecx
  loc_00DEE3E6: jz 00DEE412h
  loc_00DEE3E8: cmp [ecx], 0001h
  loc_00DEE3EC: jnz 00DEE412h
  loc_00DEE3EE: mov edi, var_14
  loc_00DEE3F1: mov edx, [ecx+00000014h]
  loc_00DEE3F4: mov eax, [ecx+00000010h]
  loc_00DEE3F7: sub edi, edx
  loc_00DEE3F9: cmp edi, eax
  loc_00DEE3FB: jb 00DEE409h
  loc_00DEE3FD: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE403: mov ecx, var_9C
  loc_00DEE409: lea eax, [edi*4]
  loc_00DEE410: jmp 00DEE41Eh
  loc_00DEE412: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE418: mov ecx, var_9C
  loc_00DEE41E: mov ecx, [ecx+0000000Ch]
  loc_00DEE421: xor edx, edx
  loc_00DEE423: mov dl, [ecx+eax]
  loc_00DEE426: mov edi, edx
  loc_00DEE428: cmp edi, 00000100h
  loc_00DEE42E: jb 00DEE436h
  loc_00DEE430: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE436: mov eax, Me
  loc_00DEE439: mov ecx, [eax+000001A8h]
  loc_00DEE43F: mov dl, [ecx+edi]
  loc_00DEE442: mov [esi], dl
  loc_00DEE444: jmp 00DEEA7Ah
  loc_00DEE449: test ecx, ecx
  loc_00DEE44B: jz 00DEE476h
  loc_00DEE44D: cmp [ecx], 0001h
  loc_00DEE451: jnz 00DEE476h
  loc_00DEE453: mov eax, var_14
  loc_00DEE456: mov edi, [ecx+00000014h]
  loc_00DEE459: mov edx, [ecx+00000010h]
  loc_00DEE45C: mov esi, eax
  loc_00DEE45E: sub esi, edi
  loc_00DEE460: cmp esi, edx
  loc_00DEE462: jb 00DEE46Dh
  loc_00DEE464: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE46A: mov eax, var_14
  loc_00DEE46D: lea edi, [esi*4]
  loc_00DEE474: jmp 00DEE481h
  loc_00DEE476: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE47C: mov edi, eax
  loc_00DEE47E: mov eax, var_14
  loc_00DEE481: mov ecx, var_40
  loc_00DEE484: test ecx, ecx
  loc_00DEE486: jz 00DEE4ACh
  loc_00DEE488: cmp [ecx], 0001h
  loc_00DEE48C: jnz 00DEE4ACh
  loc_00DEE48E: sub eax, [ecx+00000014h]
  loc_00DEE491: mov esi, eax
  loc_00DEE493: mov eax, [ecx+00000010h]
  loc_00DEE496: cmp esi, eax
  loc_00DEE498: jb 00DEE4A3h
  loc_00DEE49A: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE4A0: mov ecx, var_40
  loc_00DEE4A3: lea eax, [esi*4]
  loc_00DEE4AA: jmp 00DEE4B5h
  loc_00DEE4AC: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE4B2: mov ecx, var_40
  loc_00DEE4B5: mov edx, var_9C
  loc_00DEE4BB: mov ecx, [ecx+0000000Ch]
  loc_00DEE4BE: add ecx, eax
  loc_00DEE4C0: mov edx, [edx+0000000Ch]
  loc_00DEE4C3: add edx, edi
  loc_00DEE4C5: push edx
  loc_00DEE4C6: push ecx
  loc_00DEE4C7: push 00000004h
  loc_00DEE4C9: call [00401060h] ; __vbaCopyBytes
  loc_00DEE4CF: jmp 00DEEA7Ah
  loc_00DEE4D4: test edi, edi
  loc_00DEE4D6: jle 00DEEA7Ah
  loc_00DEE4DC: mov eax, var_3C
  loc_00DEE4DF: cmp eax, 00000001h
  loc_00DEE4E2: jnz 00DEE6E3h
  loc_00DEE4E8: mov ecx, var_9C
  loc_00DEE4EE: test ecx, ecx
  loc_00DEE4F0: jz 00DEE523h
  loc_00DEE4F2: cmp [ecx], ax
  loc_00DEE4F5: jnz 00DEE523h
  loc_00DEE4F7: mov eax, var_14
  loc_00DEE4FA: mov edx, [ecx+00000014h]
  loc_00DEE4FD: sub eax, edx
  loc_00DEE4FF: mov edx, [ecx+00000010h]
  loc_00DEE502: cmp eax, edx
  loc_00DEE504: mov var_E4, eax
  loc_00DEE50A: jb 00DEE51Eh
  loc_00DEE50C: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE512: mov ecx, var_9C
  loc_00DEE518: mov eax, var_E4
  loc_00DEE51E: shl eax, 02h
  loc_00DEE521: jmp 00DEE52Fh
  loc_00DEE523: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE529: mov ecx, var_9C
  loc_00DEE52F: mov edx, [ecx+0000000Ch]
  loc_00DEE532: xor ecx, ecx
  loc_00DEE534: mov cl, [edx+eax+00000002h]
  loc_00DEE538: mov eax, ecx
  loc_00DEE53A: cmp eax, 00000100h
  loc_00DEE53F: mov var_E8, eax
  loc_00DEE545: jb 00DEE54Dh
  loc_00DEE547: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE54D: mov edx, Me
  loc_00DEE550: xor ecx, ecx
  loc_00DEE552: mov eax, [edx+0000018Ch]
  loc_00DEE558: mov edx, var_E8
  loc_00DEE55E: mov cl, [eax+edx]
  loc_00DEE561: mov eax, ecx
  loc_00DEE563: imul eax, edi
  loc_00DEE566: jo 00DEECCEh
  loc_00DEE56C: xor ecx, ecx
  loc_00DEE56E: mov cl, [esi+00000002h]
  loc_00DEE571: imul ecx, ebx
  loc_00DEE574: jo 00DEECCEh
  loc_00DEE57A: add eax, ecx
  loc_00DEE57C: jo 00DEECCEh
  loc_00DEE582: cdq
  loc_00DEE583: and edx, 000000FFh
  loc_00DEE589: add eax, edx
  loc_00DEE58B: mov ecx, eax
  loc_00DEE58D: sar ecx, 08h
  loc_00DEE590: call [00401158h] ; __vbaUI1I4
  loc_00DEE596: mov [esi+00000002h], al
  loc_00DEE599: mov ecx, var_9C
  loc_00DEE59F: test ecx, ecx
  loc_00DEE5A1: jz 00DEE5D5h
  loc_00DEE5A3: cmp [ecx], 0001h
  loc_00DEE5A7: jnz 00DEE5D5h
  loc_00DEE5A9: mov eax, var_14
  loc_00DEE5AC: mov edx, [ecx+00000014h]
  loc_00DEE5AF: sub eax, edx
  loc_00DEE5B1: mov edx, [ecx+00000010h]
  loc_00DEE5B4: cmp eax, edx
  loc_00DEE5B6: mov var_E4, eax
  loc_00DEE5BC: jb 00DEE5D0h
  loc_00DEE5BE: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE5C4: mov ecx, var_9C
  loc_00DEE5CA: mov eax, var_E4
  loc_00DEE5D0: shl eax, 02h
  loc_00DEE5D3: jmp 00DEE5E1h
  loc_00DEE5D5: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE5DB: mov ecx, var_9C
  loc_00DEE5E1: mov edx, [ecx+0000000Ch]
  loc_00DEE5E4: xor ecx, ecx
  loc_00DEE5E6: mov cl, [edx+eax+00000001h]
  loc_00DEE5EA: mov eax, ecx
  loc_00DEE5EC: cmp eax, 00000100h
  loc_00DEE5F1: mov var_E8, eax
  loc_00DEE5F7: jb 00DEE5FFh
  loc_00DEE5F9: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE5FF: mov edx, Me
  loc_00DEE602: xor ecx, ecx
  loc_00DEE604: mov eax, [edx+0000018Ch]
  loc_00DEE60A: mov edx, var_E8
  loc_00DEE610: mov cl, [eax+edx]
  loc_00DEE613: mov eax, ecx
  loc_00DEE615: imul eax, edi
  loc_00DEE618: jo 00DEECCEh
  loc_00DEE61E: xor ecx, ecx
  loc_00DEE620: mov cl, [esi+00000001h]
  loc_00DEE623: imul ecx, ebx
  loc_00DEE626: jo 00DEECCEh
  loc_00DEE62C: add eax, ecx
  loc_00DEE62E: jo 00DEECCEh
  loc_00DEE634: cdq
  loc_00DEE635: and edx, 000000FFh
  loc_00DEE63B: add eax, edx
  loc_00DEE63D: mov ecx, eax
  loc_00DEE63F: sar ecx, 08h
  loc_00DEE642: call [00401158h] ; __vbaUI1I4
  loc_00DEE648: mov [esi+00000001h], al
  loc_00DEE64B: mov ecx, var_9C
  loc_00DEE651: test ecx, ecx
  loc_00DEE653: jz 00DEE687h
  loc_00DEE655: cmp [ecx], 0001h
  loc_00DEE659: jnz 00DEE687h
  loc_00DEE65B: mov eax, var_14
  loc_00DEE65E: mov edx, [ecx+00000014h]
  loc_00DEE661: sub eax, edx
  loc_00DEE663: mov edx, [ecx+00000010h]
  loc_00DEE666: cmp eax, edx
  loc_00DEE668: mov var_E4, eax
  loc_00DEE66E: jb 00DEE682h
  loc_00DEE670: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE676: mov ecx, var_9C
  loc_00DEE67C: mov eax, var_E4
  loc_00DEE682: shl eax, 02h
  loc_00DEE685: jmp 00DEE693h
  loc_00DEE687: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE68D: mov ecx, var_9C
  loc_00DEE693: mov edx, [ecx+0000000Ch]
  loc_00DEE696: xor ecx, ecx
  loc_00DEE698: mov cl, [edx+eax]
  loc_00DEE69B: mov eax, ecx
  loc_00DEE69D: cmp eax, 00000100h
  loc_00DEE6A2: mov var_E8, eax
  loc_00DEE6A8: jb 00DEE6B0h
  loc_00DEE6AA: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE6B0: mov edx, Me
  loc_00DEE6B3: xor ecx, ecx
  loc_00DEE6B5: mov eax, [edx+0000018Ch]
  loc_00DEE6BB: mov edx, var_E8
  loc_00DEE6C1: mov cl, [eax+edx]
  loc_00DEE6C4: mov eax, ecx
  loc_00DEE6C6: imul eax, edi
  loc_00DEE6C9: jo 00DEECCEh
  loc_00DEE6CF: xor ecx, ecx
  loc_00DEE6D1: mov cl, [esi]
  loc_00DEE6D3: imul ecx, ebx
  loc_00DEE6D6: jo 00DEECCEh
  loc_00DEE6DC: add eax, ecx
  loc_00DEE6DE: jmp 00DEEA5Eh
  loc_00DEE6E3: mov ecx, var_9C
  loc_00DEE6E9: cmp eax, 00000002h
  loc_00DEE6EC: jnz 00DEE8E8h
  loc_00DEE6F2: test ecx, ecx
  loc_00DEE6F4: jz 00DEE728h
  loc_00DEE6F6: cmp [ecx], 0001h
  loc_00DEE6FA: jnz 00DEE728h
  loc_00DEE6FC: mov eax, var_14
  loc_00DEE6FF: mov edx, [ecx+00000014h]
  loc_00DEE702: sub eax, edx
  loc_00DEE704: mov edx, [ecx+00000010h]
  loc_00DEE707: cmp eax, edx
  loc_00DEE709: mov var_E4, eax
  loc_00DEE70F: jb 00DEE723h
  loc_00DEE711: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE717: mov ecx, var_9C
  loc_00DEE71D: mov eax, var_E4
  loc_00DEE723: shl eax, 02h
  loc_00DEE726: jmp 00DEE734h
  loc_00DEE728: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE72E: mov ecx, var_9C
  loc_00DEE734: mov edx, [ecx+0000000Ch]
  loc_00DEE737: xor ecx, ecx
  loc_00DEE739: mov cl, [edx+eax+00000002h]
  loc_00DEE73D: mov eax, ecx
  loc_00DEE73F: cmp eax, 00000100h
  loc_00DEE744: mov var_E8, eax
  loc_00DEE74A: jb 00DEE752h
  loc_00DEE74C: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE752: mov edx, Me
  loc_00DEE755: xor ecx, ecx
  loc_00DEE757: mov eax, [edx+000001A8h]
  loc_00DEE75D: mov edx, var_E8
  loc_00DEE763: mov cl, [eax+edx]
  loc_00DEE766: mov eax, ecx
  loc_00DEE768: imul eax, edi
  loc_00DEE76B: jo 00DEECCEh
  loc_00DEE771: xor ecx, ecx
  loc_00DEE773: mov cl, [esi+00000002h]
  loc_00DEE776: imul ecx, ebx
  loc_00DEE779: jo 00DEECCEh
  loc_00DEE77F: add eax, ecx
  loc_00DEE781: jo 00DEECCEh
  loc_00DEE787: cdq
  loc_00DEE788: and edx, 000000FFh
  loc_00DEE78E: add eax, edx
  loc_00DEE790: mov ecx, eax
  loc_00DEE792: sar ecx, 08h
  loc_00DEE795: call [00401158h] ; __vbaUI1I4
  loc_00DEE79B: mov [esi+00000002h], al
  loc_00DEE79E: mov ecx, var_9C
  loc_00DEE7A4: test ecx, ecx
  loc_00DEE7A6: jz 00DEE7DAh
  loc_00DEE7A8: cmp [ecx], 0001h
  loc_00DEE7AC: jnz 00DEE7DAh
  loc_00DEE7AE: mov eax, var_14
  loc_00DEE7B1: mov edx, [ecx+00000014h]
  loc_00DEE7B4: sub eax, edx
  loc_00DEE7B6: mov edx, [ecx+00000010h]
  loc_00DEE7B9: cmp eax, edx
  loc_00DEE7BB: mov var_E4, eax
  loc_00DEE7C1: jb 00DEE7D5h
  loc_00DEE7C3: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE7C9: mov ecx, var_9C
  loc_00DEE7CF: mov eax, var_E4
  loc_00DEE7D5: shl eax, 02h
  loc_00DEE7D8: jmp 00DEE7E6h
  loc_00DEE7DA: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE7E0: mov ecx, var_9C
  loc_00DEE7E6: mov edx, [ecx+0000000Ch]
  loc_00DEE7E9: xor ecx, ecx
  loc_00DEE7EB: mov cl, [edx+eax+00000001h]
  loc_00DEE7EF: mov eax, ecx
  loc_00DEE7F1: cmp eax, 00000100h
  loc_00DEE7F6: mov var_E8, eax
  loc_00DEE7FC: jb 00DEE804h
  loc_00DEE7FE: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE804: mov edx, Me
  loc_00DEE807: xor ecx, ecx
  loc_00DEE809: mov eax, [edx+000001A8h]
  loc_00DEE80F: mov edx, var_E8
  loc_00DEE815: mov cl, [eax+edx]
  loc_00DEE818: mov eax, ecx
  loc_00DEE81A: imul eax, edi
  loc_00DEE81D: jo 00DEECCEh
  loc_00DEE823: xor ecx, ecx
  loc_00DEE825: mov cl, [esi+00000001h]
  loc_00DEE828: imul ecx, ebx
  loc_00DEE82B: jo 00DEECCEh
  loc_00DEE831: add eax, ecx
  loc_00DEE833: jo 00DEECCEh
  loc_00DEE839: cdq
  loc_00DEE83A: and edx, 000000FFh
  loc_00DEE840: add eax, edx
  loc_00DEE842: mov ecx, eax
  loc_00DEE844: sar ecx, 08h
  loc_00DEE847: call [00401158h] ; __vbaUI1I4
  loc_00DEE84D: mov [esi+00000001h], al
  loc_00DEE850: mov ecx, var_9C
  loc_00DEE856: test ecx, ecx
  loc_00DEE858: jz 00DEE88Ch
  loc_00DEE85A: cmp [ecx], 0001h
  loc_00DEE85E: jnz 00DEE88Ch
  loc_00DEE860: mov eax, var_14
  loc_00DEE863: mov edx, [ecx+00000014h]
  loc_00DEE866: sub eax, edx
  loc_00DEE868: mov edx, [ecx+00000010h]
  loc_00DEE86B: cmp eax, edx
  loc_00DEE86D: mov var_E4, eax
  loc_00DEE873: jb 00DEE887h
  loc_00DEE875: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE87B: mov ecx, var_9C
  loc_00DEE881: mov eax, var_E4
  loc_00DEE887: shl eax, 02h
  loc_00DEE88A: jmp 00DEE898h
  loc_00DEE88C: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE892: mov ecx, var_9C
  loc_00DEE898: mov edx, [ecx+0000000Ch]
  loc_00DEE89B: xor ecx, ecx
  loc_00DEE89D: mov cl, [edx+eax]
  loc_00DEE8A0: mov eax, ecx
  loc_00DEE8A2: cmp eax, 00000100h
  loc_00DEE8A7: mov var_E8, eax
  loc_00DEE8AD: jb 00DEE8B5h
  loc_00DEE8AF: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE8B5: mov edx, Me
  loc_00DEE8B8: xor ecx, ecx
  loc_00DEE8BA: mov eax, [edx+000001A8h]
  loc_00DEE8C0: mov edx, var_E8
  loc_00DEE8C6: mov cl, [eax+edx]
  loc_00DEE8C9: mov eax, ecx
  loc_00DEE8CB: imul eax, edi
  loc_00DEE8CE: jo 00DEECCEh
  loc_00DEE8D4: xor ecx, ecx
  loc_00DEE8D6: mov cl, [esi]
  loc_00DEE8D8: imul ecx, ebx
  loc_00DEE8DB: jo 00DEECCEh
  loc_00DEE8E1: add eax, ecx
  loc_00DEE8E3: jmp 00DEEA5Eh
  loc_00DEE8E8: test ecx, ecx
  loc_00DEE8EA: jz 00DEE91Eh
  loc_00DEE8EC: cmp [ecx], 0001h
  loc_00DEE8F0: jnz 00DEE91Eh
  loc_00DEE8F2: mov eax, var_14
  loc_00DEE8F5: mov edx, [ecx+00000014h]
  loc_00DEE8F8: sub eax, edx
  loc_00DEE8FA: mov edx, [ecx+00000010h]
  loc_00DEE8FD: cmp eax, edx
  loc_00DEE8FF: mov var_E4, eax
  loc_00DEE905: jb 00DEE919h
  loc_00DEE907: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE90D: mov ecx, var_9C
  loc_00DEE913: mov eax, var_E4
  loc_00DEE919: shl eax, 02h
  loc_00DEE91C: jmp 00DEE92Ah
  loc_00DEE91E: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE924: mov ecx, var_9C
  loc_00DEE92A: mov edx, [ecx+0000000Ch]
  loc_00DEE92D: xor ecx, ecx
  loc_00DEE92F: mov cl, [edx+eax+00000002h]
  loc_00DEE933: mov eax, ecx
  loc_00DEE935: imul eax, edi
  loc_00DEE938: jo 00DEECCEh
  loc_00DEE93E: xor edx, edx
  loc_00DEE940: mov dl, [esi+00000002h]
  loc_00DEE943: imul edx, ebx
  loc_00DEE946: jo 00DEECCEh
  loc_00DEE94C: add eax, edx
  loc_00DEE94E: jo 00DEECCEh
  loc_00DEE954: cdq
  loc_00DEE955: and edx, 000000FFh
  loc_00DEE95B: add eax, edx
  loc_00DEE95D: mov ecx, eax
  loc_00DEE95F: sar ecx, 08h
  loc_00DEE962: call [00401158h] ; __vbaUI1I4
  loc_00DEE968: mov [esi+00000002h], al
  loc_00DEE96B: mov ecx, var_9C
  loc_00DEE971: test ecx, ecx
  loc_00DEE973: jz 00DEE9A7h
  loc_00DEE975: cmp [ecx], 0001h
  loc_00DEE979: jnz 00DEE9A7h
  loc_00DEE97B: mov eax, var_14
  loc_00DEE97E: mov edx, [ecx+00000014h]
  loc_00DEE981: sub eax, edx
  loc_00DEE983: mov edx, [ecx+00000010h]
  loc_00DEE986: cmp eax, edx
  loc_00DEE988: mov var_E4, eax
  loc_00DEE98E: jb 00DEE9A2h
  loc_00DEE990: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE996: mov ecx, var_9C
  loc_00DEE99C: mov eax, var_E4
  loc_00DEE9A2: shl eax, 02h
  loc_00DEE9A5: jmp 00DEE9B3h
  loc_00DEE9A7: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEE9AD: mov ecx, var_9C
  loc_00DEE9B3: mov ecx, [ecx+0000000Ch]
  loc_00DEE9B6: xor edx, edx
  loc_00DEE9B8: mov dl, [ecx+eax+00000001h]
  loc_00DEE9BC: mov eax, edx
  loc_00DEE9BE: imul eax, edi
  loc_00DEE9C1: jo 00DEECCEh
  loc_00DEE9C7: xor ecx, ecx
  loc_00DEE9C9: mov cl, [esi+00000001h]
  loc_00DEE9CC: imul ecx, ebx
  loc_00DEE9CF: jo 00DEECCEh
  loc_00DEE9D5: add eax, ecx
  loc_00DEE9D7: jo 00DEECCEh
  loc_00DEE9DD: cdq
  loc_00DEE9DE: and edx, 000000FFh
  loc_00DEE9E4: add eax, edx
  loc_00DEE9E6: mov ecx, eax
  loc_00DEE9E8: sar ecx, 08h
  loc_00DEE9EB: call [00401158h] ; __vbaUI1I4
  loc_00DEE9F1: mov [esi+00000001h], al
  loc_00DEE9F4: mov ecx, var_9C
  loc_00DEE9FA: test ecx, ecx
  loc_00DEE9FC: jz 00DEEA30h
  loc_00DEE9FE: cmp [ecx], 0001h
  loc_00DEEA02: jnz 00DEEA30h
  loc_00DEEA04: mov eax, var_14
  loc_00DEEA07: mov edx, [ecx+00000014h]
  loc_00DEEA0A: sub eax, edx
  loc_00DEEA0C: mov edx, [ecx+00000010h]
  loc_00DEEA0F: cmp eax, edx
  loc_00DEEA11: mov var_E4, eax
  loc_00DEEA17: jb 00DEEA2Bh
  loc_00DEEA19: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEEA1F: mov ecx, var_9C
  loc_00DEEA25: mov eax, var_E4
  loc_00DEEA2B: shl eax, 02h
  loc_00DEEA2E: jmp 00DEEA3Ch
  loc_00DEEA30: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEEA36: mov ecx, var_9C
  loc_00DEEA3C: mov edx, [ecx+0000000Ch]
  loc_00DEEA3F: xor ecx, ecx
  loc_00DEEA41: mov cl, [edx+eax]
  loc_00DEEA44: mov eax, ecx
  loc_00DEEA46: imul eax, edi
  loc_00DEEA49: jo 00DEECCEh
  loc_00DEEA4F: xor edx, edx
  loc_00DEEA51: mov dl, [esi]
  loc_00DEEA53: imul edx, ebx
  loc_00DEEA56: jo 00DEECCEh
  loc_00DEEA5C: add eax, edx
  loc_00DEEA5E: jo 00DEECCEh
  loc_00DEEA64: cdq
  loc_00DEEA65: and edx, 000000FFh
  loc_00DEEA6B: add eax, edx
  loc_00DEEA6D: mov ecx, eax
  loc_00DEEA6F: sar ecx, 08h
  loc_00DEEA72: call [00401158h] ; __vbaUI1I4
  loc_00DEEA78: mov [esi], al
  loc_00DEEA7A: lea eax, var_F0
  loc_00DEEA80: push eax
  loc_00DEEA81: call [00401244h] ; __vbaAryUnlock
  loc_00DEEA87: mov ecx, var_70
  loc_00DEEA8A: mov edx, var_A0
  loc_00DEEA90: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DEEA96: mov ebx, arg_1C
  loc_00DEEA99: mov esi, arg_18
  loc_00DEEA9C: mov eax, 00000001h
  loc_00DEEAA1: add eax, ecx
  loc_00DEEAA3: mov ecx, var_34
  loc_00DEEAA6: jo 00DEECCEh
  loc_00DEEAAC: mov var_70, eax
  loc_00DEEAAF: mov eax, var_94
  loc_00DEEAB5: jmp 00DEDC73h
  loc_00DEEABA: mov eax, 00000001h
  loc_00DEEABF: add eax, edx
  loc_00DEEAC1: jo 00DEECCEh
  loc_00DEEAC7: mov edx, eax
  loc_00DEEAC9: mov eax, var_104
  loc_00DEEACF: mov var_A0, edx
  loc_00DEEAD5: jmp 00DEDC53h
  loc_00DEEADA: mov ecx, var_40
  loc_00DEEADD: lea edx, var_A4
  loc_00DEEAE3: push ecx
  loc_00DEEAE4: push edx
  loc_00DEEAE5: call [004011DCh] ; __vbaAryLock
  loc_00DEEAEB: mov ecx, var_A4
  loc_00DEEAF1: test ecx, ecx
  loc_00DEEAF3: jz 00DEEB24h
  loc_00DEEAF5: cmp [ecx], 0001h
  loc_00DEEAF9: jnz 00DEEB24h
  loc_00DEEAFB: mov eax, [ecx+00000014h]
  loc_00DEEAFE: mov edx, [ecx+00000010h]
  loc_00DEEB01: neg eax
  loc_00DEEB03: cmp eax, edx
  loc_00DEEB05: mov var_E4, eax
  loc_00DEEB0B: jb 00DEEB1Fh
  loc_00DEEB0D: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEEB13: mov ecx, var_A4
  loc_00DEEB19: mov eax, var_E4
  loc_00DEEB1F: shl eax, 02h
  loc_00DEEB22: jmp 00DEEB30h
  loc_00DEEB24: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEEB2A: mov ecx, var_A4
  loc_00DEEB30: mov ecx, [ecx+0000000Ch]
  loc_00DEEB33: lea edx, var_6C
  loc_00DEEB36: push 00000000h
  loc_00DEEB38: add ecx, eax
  loc_00DEEB3A: mov eax, arg_10
  loc_00DEEB3D: push edx
  loc_00DEEB3E: mov edx, arg_14
  loc_00DEEB41: push ecx
  loc_00DEEB42: mov ecx, arg_C
  loc_00DEEB45: push ebx
  loc_00DEEB46: push 00000000h
  loc_00DEEB48: push 00000000h
  loc_00DEEB4A: push 00000000h
  loc_00DEEB4C: push ebx
  loc_00DEEB4D: push esi
  loc_00DEEB4E: push edx
  loc_00DEEB4F: push eax
  loc_00DEEB50: push ecx
  loc_00DEEB51: call 006DD340h ; SetDIBitsToDevice(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v, %x10v, %x11v, %x12v)
  loc_00DEEB56: call edi
  loc_00DEEB58: lea edx, var_A4
  loc_00DEEB5E: push edx
  loc_00DEEB5F: call [00401244h] ; __vbaAryUnlock
  loc_00DEEB65: mov esi, [004010E8h] ; __vbaErase
  loc_00DEEB6B: lea eax, var_9C
  loc_00DEEB71: push eax
  loc_00DEEB72: push 00000000h
  loc_00DEEB74: call __vbaErase
  loc_00DEEB76: lea ecx, var_40
  loc_00DEEB79: push ecx
  loc_00DEEB7A: push 00000000h
  loc_00DEEB7C: call __vbaErase
  loc_00DEEB7E: mov edx, var_20
  loc_00DEEB81: mov eax, var_84
  loc_00DEEB87: push edx
  loc_00DEEB88: push eax
  loc_00DEEB89: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEEB8E: mov var_DC, eax
  loc_00DEEB94: call edi
  loc_00DEEB96: mov ecx, var_DC
  loc_00DEEB9C: push ecx
  loc_00DEEB9D: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEEBA2: call edi
  loc_00DEEBA4: mov edx, var_30
  loc_00DEEBA7: mov eax, var_38
  loc_00DEEBAA: push edx
  loc_00DEEBAB: push eax
  loc_00DEEBAC: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEEBB1: mov var_DC, eax
  loc_00DEEBB7: call edi
  loc_00DEEBB9: mov ecx, var_DC
  loc_00DEEBBF: push ecx
  loc_00DEEBC0: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEEBC5: call edi
  loc_00DEEBC7: mov edx, var_24
  loc_00DEEBCA: push 00000000h
  loc_00DEEBCC: push 00000003h
  loc_00DEEBCE: lea eax, var_B4
  loc_00DEEBD4: push edx
  loc_00DEEBD5: push eax
  loc_00DEEBD6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DEEBDC: add esp, 00000010h
  loc_00DEEBDF: push eax
  loc_00DEEBE0: call [0040118Ch] ; __vbaI2Var
  loc_00DEEBE6: xor ecx, ecx
  loc_00DEEBE8: cmp ax, 0003h
  loc_00DEEBEC: setz cl
  loc_00DEEBEF: neg ecx
  loc_00DEEBF1: mov esi, ecx
  loc_00DEEBF3: lea ecx, var_B4
  loc_00DEEBF9: call [00401024h] ; __vbaFreeVar
  loc_00DEEBFF: test si, si
  loc_00DEEC02: jz 00DEEC2Dh
  loc_00DEEC04: mov edx, var_98
  loc_00DEEC0A: mov eax, var_E0
  loc_00DEEC10: push edx
  loc_00DEEC11: push eax
  loc_00DEEC12: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEEC17: mov var_DC, eax
  loc_00DEEC1D: call edi
  loc_00DEEC1F: mov ecx, var_DC
  loc_00DEEC25: push ecx
  loc_00DEEC26: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEEC2B: call edi
  loc_00DEEC2D: mov edx, var_84
  loc_00DEEC33: push edx
  loc_00DEEC34: call 006DD4DCh ; DeleteDC(%x1v)
  loc_00DEEC39: call edi
  loc_00DEEC3B: mov eax, var_38
  loc_00DEEC3E: push eax
  loc_00DEEC3F: call 006DD4DCh ; DeleteDC(%x1v)
  loc_00DEEC44: call edi
  loc_00DEEC46: mov ecx, var_98
  loc_00DEEC4C: push ecx
  loc_00DEEC4D: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEEC52: call edi
  loc_00DEEC54: mov edx, var_E0
  loc_00DEEC5A: push edx
  loc_00DEEC5B: call 006DD4DCh ; DeleteDC(%x1v)
  loc_00DEEC60: call edi
  loc_00DEEC62: fwait
  loc_00DEEC63: push 00DEECB4h
  loc_00DEEC68: jmp 00DEEC84h
  loc_00DEEC6A: lea eax, var_A4
  loc_00DEEC70: push eax
  loc_00DEEC71: call [00401244h] ; __vbaAryUnlock
  loc_00DEEC77: lea ecx, var_B4
  loc_00DEEC7D: call [00401024h] ; __vbaFreeVar
  loc_00DEEC83: ret
  loc_00DEEC84: lea ecx, var_F0
  loc_00DEEC8A: push ecx
  loc_00DEEC8B: call [00401244h] ; __vbaAryUnlock
  loc_00DEEC91: lea ecx, var_24
  loc_00DEEC94: call [00401254h] ; __vbaFreeObj
  loc_00DEEC9A: mov esi, [00401090h] ; __vbaAryDestruct
  loc_00DEECA0: lea edx, var_40
  loc_00DEECA3: push edx
  loc_00DEECA4: push 00000000h
  loc_00DEECA6: call __vbaAryDestruct
  loc_00DEECA8: lea eax, var_9C
  loc_00DEECAE: push eax
  loc_00DEECAF: push 00000000h
  loc_00DEECB1: call __vbaAryDestruct
  loc_00DEECB3: ret
  loc_00DEECB4: mov ecx, var_10
  loc_00DEECB7: pop edi
  loc_00DEECB8: pop esi
  loc_00DEECB9: xor eax, eax
  loc_00DEECBB: mov fs:[00000000h], ecx
  loc_00DEECC2: pop ebx
  loc_00DEECC3: mov esp, ebp
  loc_00DEECC5: pop ebp
  loc_00DEECC6: retn 0024h
End Function

Private Sub Proc_2_107_DEECE0
  loc_00DEECE0: mov ecx, [esp+00000008h]
  loc_00DEECE4: mov eax, 80808081h
  loc_00DEECE9: and ecx, 000000FFh
  loc_00DEECEF: imul ecx, ecx, 00000125h
  loc_00DEECF5: jo 00DEED36h
  loc_00DEECF7: imul ecx
  loc_00DEECF9: add edx, ecx
  loc_00DEECFB: sar edx, 07h
  loc_00DEECFE: mov eax, edx
  loc_00DEED00: shr eax, 1Fh
  loc_00DEED03: add edx, eax
  loc_00DEED05: mov ecx, edx
  loc_00DEED07: cmp ecx, 000000FFh
  loc_00DEED0D: jle 00DEED25h
  loc_00DEED0F: mov ecx, 000000FFh
  loc_00DEED14: call [00401144h] ; __vbaUI1I2
  loc_00DEED1A: mov ecx, Proc_2_107_DEECE0
  loc_00DEED1E: mov [ecx], al
  loc_00DEED20: xor eax, eax
  loc_00DEED22: retn 000Ch
End Sub

Private Sub Proc_2_108_DEED40
  loc_00DEED40: push esi
  loc_00DEED41: mov esi, Proc_2_108_DEED40
  loc_00DEED45: and esi, 000000FFh
  loc_00DEED4B: mov eax, 80808081h
  loc_00DEED50: imul esi, esi, 000000D9h
  loc_00DEED56: jo 00DEED7Ah
  loc_00DEED58: imul esi
  loc_00DEED5A: mov ecx, edx
  loc_00DEED5C: add ecx, esi
  loc_00DEED5E: sar ecx, 07h
  loc_00DEED61: mov eax, ecx
  loc_00DEED63: shr eax, 1Fh
  loc_00DEED66: add ecx, eax
  loc_00DEED68: call [00401158h] ; __vbaUI1I4
  loc_00DEED6E: mov ecx, Me
  loc_00DEED72: pop esi
  loc_00DEED73: mov [ecx], al
  loc_00DEED75: xor eax, eax
  loc_00DEED77: retn 000Ch
End Sub

Private Function Proc_2_109_DEED80(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24) 'DEED80
  loc_00DEED80: push ebp
  loc_00DEED81: mov ebp, esp
  loc_00DEED83: sub esp, 00000008h
  loc_00DEED86: push 00402836h ; __vbaExceptHandler
  loc_00DEED8B: mov eax, fs:[00000000h]
  loc_00DEED91: push eax
  loc_00DEED92: mov fs:[00000000h], esp
  loc_00DEED99: sub esp, 000000F4h
  loc_00DEED9F: push ebx
  loc_00DEEDA0: push esi
  loc_00DEEDA1: push edi
  loc_00DEEDA2: mov var_8, esp
  loc_00DEEDA5: mov var_4, 00401380h
  loc_00DEEDAC: mov ecx, 0000000Ah
  loc_00DEEDB1: xor eax, eax
  loc_00DEEDB3: lea edi, var_70
  loc_00DEEDB6: xor edx, edx
  loc_00DEEDB8: repz stosd
  loc_00DEEDBA: mov ecx, arg_14
  loc_00DEEDBD: mov eax, 00000001h
  loc_00DEEDC2: cmp ecx, eax
  loc_00DEEDC4: mov var_24, edx
  loc_00DEEDC7: mov var_40, edx
  loc_00DEEDCA: mov var_44, edx
  loc_00DEEDCD: mov var_80, edx
  loc_00DEEDD0: mov var_84, edx
  loc_00DEEDD6: mov var_88, edx
  loc_00DEEDDC: jl 00DEF684h
  loc_00DEEDE2: cmp arg_18, eax
  loc_00DEEDE5: jl 00DEF684h
  loc_00DEEDEB: mov eax, arg_1C
  loc_00DEEDEE: and eax, 00FFFFFFh
  loc_00DEEDF3: mov edi, eax
  loc_00DEEDF5: and edi, 800000FFh
  loc_00DEEDFB: jns 00DEEE05h
  loc_00DEEDFD: dec edi
  loc_00DEEDFE: or edi, FFFFFF00h
  loc_00DEEE04: inc edi
  loc_00DEEE05: cdq
  loc_00DEEE06: and edx, 000000FFh
  loc_00DEEE0C: add eax, edx
  loc_00DEEE0E: sar eax, 08h
  loc_00DEEE11: mov ebx, eax
  loc_00DEEE13: and ebx, 800000FFh
  loc_00DEEE19: jns 00DEEE23h
  loc_00DEEE1B: dec ebx
  loc_00DEEE1C: or ebx, FFFFFF00h
  loc_00DEEE22: inc ebx
  loc_00DEEE23: cdq
  loc_00DEEE24: and edx, 000000FFh
  loc_00DEEE2A: mov var_48, ebx
  loc_00DEEE2D: add eax, edx
  loc_00DEEE2F: sar eax, 08h
  loc_00DEEE32: mov esi, eax
  loc_00DEEE34: and esi, 800000FFh
  loc_00DEEE3A: jns 00DEEE44h
  loc_00DEEE3C: dec esi
  loc_00DEEE3D: or esi, FFFFFF00h
  loc_00DEEE43: inc esi
  loc_00DEEE44: mov eax, arg_20
  loc_00DEEE47: mov var_1C, esi
  loc_00DEEE4A: and eax, 00FFFFFFh
  loc_00DEEE4F: mov ecx, eax
  loc_00DEEE51: and ecx, 800000FFh
  loc_00DEEE57: jns 00DEEE61h
  loc_00DEEE59: dec ecx
  loc_00DEEE5A: or ecx, FFFFFF00h
  loc_00DEEE60: inc ecx
  loc_00DEEE61: cdq
  loc_00DEEE62: and edx, 000000FFh
  loc_00DEEE68: mov var_34, ecx
  loc_00DEEE6B: add eax, edx
  loc_00DEEE6D: sar eax, 08h
  loc_00DEEE70: mov edx, eax
  loc_00DEEE72: and edx, 800000FFh
  loc_00DEEE78: jns 00DEEE82h
  loc_00DEEE7A: dec edx
  loc_00DEEE7B: or edx, FFFFFF00h
  loc_00DEEE81: inc edx
  loc_00DEEE82: mov var_74, edx
  loc_00DEEE85: cdq
  loc_00DEEE86: and edx, 000000FFh
  loc_00DEEE8C: add eax, edx
  loc_00DEEE8E: sar eax, 08h
  loc_00DEEE91: mov edx, eax
  loc_00DEEE93: and edx, 800000FFh
  loc_00DEEE99: jns 00DEEEA3h
  loc_00DEEE9B: dec edx
  loc_00DEEE9C: or edx, FFFFFF00h
  loc_00DEEEA2: inc edx
  loc_00DEEEA3: mov eax, ecx
  loc_00DEEEA5: mov var_30, edx
  loc_00DEEEA8: sub eax, edi
  loc_00DEEEAA: jo 00DEF6D8h
  loc_00DEEEB0: mov var_7C, eax
  loc_00DEEEB3: mov eax, var_74
  loc_00DEEEB6: sub eax, ebx
  loc_00DEEEB8: jo 00DEF6D8h
  loc_00DEEEBE: mov var_18, eax
  loc_00DEEEC1: mov eax, edx
  loc_00DEEEC3: sub eax, esi
  loc_00DEEEC5: jo 00DEF6D8h
  loc_00DEEECB: mov var_78, eax
  loc_00DEEECE: mov eax, arg_24
  loc_00DEEED1: sub eax, 00000000h
  loc_00DEEED4: jz 00DEEF14h
  loc_00DEEED6: dec eax
  loc_00DEEED7: push 00000000h
  loc_00DEEED9: jz 00DEEEFDh
  loc_00DEEEDB: mov eax, arg_14
  loc_00DEEEDE: mov ecx, arg_18
  loc_00DEEEE1: add eax, ecx
  loc_00DEEEE3: lea ecx, var_24
  loc_00DEEEE6: jo 00DEF6D8h
  loc_00DEEEEC: sub eax, 00000002h
  loc_00DEEEEF: jo 00DEF6D8h
  loc_00DEEEF5: push eax
  loc_00DEEEF6: push 00000001h
  loc_00DEEEF8: push 00000003h
  loc_00DEEEFA: push ecx
  loc_00DEEEFB: jmp 00DEEF2Bh
  loc_00DEEEFD: mov edx, arg_18
  loc_00DEEF00: lea eax, var_24
  loc_00DEEF03: sub edx, 00000001h
  loc_00DEEF06: jo 00DEF6D8h
  loc_00DEEF0C: push edx
  loc_00DEEF0D: push 00000001h
  loc_00DEEF0F: push 00000003h
  loc_00DEEF11: push eax
  loc_00DEEF12: jmp 00DEEF2Bh
  loc_00DEEF14: mov ecx, arg_14
  loc_00DEEF17: push 00000000h
  loc_00DEEF19: sub ecx, 00000001h
  loc_00DEEF1C: lea edx, var_24
  loc_00DEEF1F: jo 00DEF6D8h
  loc_00DEEF25: push ecx
  loc_00DEEF26: push 00000001h
  loc_00DEEF28: push 00000003h
  loc_00DEEF2A: push edx
  loc_00DEEF2B: push 00000004h
  loc_00DEEF2D: push 00000080h
  loc_00DEEF32: call [00401138h] ; __vbaRedim
  loc_00DEEF38: mov eax, var_24
  loc_00DEEF3B: add esp, 0000001Ch
  loc_00DEEF3E: push eax
  loc_00DEEF3F: push 00000001h
  loc_00DEEF41: call [0040117Ch] ; __vbaUbound
  loc_00DEEF47: test eax, eax
  loc_00DEEF49: mov var_38, eax
  loc_00DEEF4C: jnz 00DEF0F5h
  loc_00DEEF52: mov ecx, var_24
  loc_00DEEF55: test ecx, ecx
  loc_00DEEF57: jz 00DEEF8Fh
  loc_00DEEF59: cmp [ecx], 0001h
  loc_00DEEF5D: jnz 00DEEF8Fh
  loc_00DEEF5F: mov eax, [ecx+00000014h]
  loc_00DEEF62: mov edx, [ecx+00000010h]
  loc_00DEEF65: neg eax
  loc_00DEEF67: cmp eax, edx
  loc_00DEEF69: mov var_8C, eax
  loc_00DEEF6F: jb 00DEEF80h
  loc_00DEEF71: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEEF77: mov ecx, var_24
  loc_00DEEF7A: mov eax, var_8C
  loc_00DEEF80: lea edx, [eax*4]
  loc_00DEEF87: mov var_EC, edx
  loc_00DEEF8D: jmp 00DEEF9Eh
  loc_00DEEF8F: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEEF95: mov ecx, var_24
  loc_00DEEF98: mov var_EC, eax
  loc_00DEEF9E: mov eax, var_74
  loc_00DEEFA1: cdq
  loc_00DEEFA2: sub eax, edx
  loc_00DEEFA4: mov edx, eax
  loc_00DEEFA6: mov eax, ebx
  loc_00DEEFA8: sar edx, 01h
  loc_00DEEFAA: mov ebx, edx
  loc_00DEEFAC: cdq
  loc_00DEEFAD: sub eax, edx
  loc_00DEEFAF: sar eax, 01h
  loc_00DEEFB1: add ebx, eax
  loc_00DEEFB3: mov eax, var_30
  loc_00DEEFB6: jo 00DEF6D8h
  loc_00DEEFBC: imul ebx, ebx, 00000100h
  loc_00DEEFC2: cdq
  loc_00DEEFC3: jo 00DEF6D8h
  loc_00DEEFC9: sub eax, edx
  loc_00DEEFCB: mov var_F4, ebx
  loc_00DEEFD1: mov ebx, eax
  loc_00DEEFD3: mov eax, esi
  loc_00DEEFD5: cdq
  loc_00DEEFD6: mov esi, var_F4
  loc_00DEEFDC: sub eax, edx
  loc_00DEEFDE: sar ebx, 01h
  loc_00DEEFE0: sar eax, 01h
  loc_00DEEFE2: add ebx, eax
  loc_00DEEFE4: mov eax, var_34
  loc_00DEEFE7: jo 00DEF6D8h
  loc_00DEEFED: add esi, ebx
  loc_00DEEFEF: cdq
  loc_00DEEFF0: jo 00DEF6D8h
  loc_00DEEFF6: sub eax, edx
  loc_00DEEFF8: mov ebx, eax
  loc_00DEEFFA: mov eax, edi
  loc_00DEEFFC: cdq
  loc_00DEEFFD: sub eax, edx
  loc_00DEEFFF: sar ebx, 01h
  loc_00DEF001: sar eax, 01h
  loc_00DEF003: add ebx, eax
  loc_00DEF005: mov eax, [ecx+0000000Ch]
  loc_00DEF008: jo 00DEF6D8h
  loc_00DEF00E: imul ebx, ebx, 00010000h
  loc_00DEF014: mov ecx, var_EC
  loc_00DEF01A: jo 00DEF6D8h
  loc_00DEF020: add esi, ebx
  loc_00DEF022: jo 00DEF6D8h
  loc_00DEF028: mov [eax+ecx], esi
  loc_00DEF02B: mov ebx, arg_14
  loc_00DEF02E: mov esi, arg_18
  loc_00DEF031: mov ecx, ebx
  loc_00DEF033: push 00000000h
  loc_00DEF035: imul ecx, esi
  loc_00DEF038: jo 00DEF6D8h
  loc_00DEF03E: sub ecx, 00000001h
  loc_00DEF041: lea edx, var_44
  loc_00DEF044: jo 00DEF6D8h
  loc_00DEF04A: push ecx
  loc_00DEF04B: push 00000001h
  loc_00DEF04D: push 00000003h
  loc_00DEF04F: push edx
  loc_00DEF050: push 00000004h
  loc_00DEF052: push 00000080h
  loc_00DEF057: call [00401138h] ; __vbaRedim
  loc_00DEF05D: mov ecx, arg_24
  loc_00DEF060: mov edi, ebx
  loc_00DEF062: add esp, 0000001Ch
  loc_00DEF065: sub edi, 00000001h
  loc_00DEF068: mov eax, esi
  loc_00DEF06A: jo 00DEF6D8h
  loc_00DEF070: sub eax, 00000001h
  loc_00DEF073: mov var_38, edi
  loc_00DEF076: jo 00DEF6D8h
  loc_00DEF07C: cmp ecx, 00000003h
  loc_00DEF07F: mov var_28, eax
  loc_00DEF082: ja 00DEF5ACh
  loc_00DEF088: jmp [ecx*4+00DEF6C8h]
  loc_00DEF08F: mov edx, var_80
  loc_00DEF092: xor ecx, ecx
  loc_00DEF094: mov var_2C, ecx
  loc_00DEF097: cmp ecx, eax
  loc_00DEF099: jg 00DEF5ACh
  loc_00DEF09F: mov ecx, edx
  loc_00DEF0A1: mov ebx, edx
  loc_00DEF0A3: add ecx, edi
  loc_00DEF0A5: jo 00DEF6D8h
  loc_00DEF0AB: mov var_B4, ecx
  loc_00DEF0B1: cmp ebx, ecx
  loc_00DEF0B3: jg 00DEF25Eh
  loc_00DEF0B9: mov eax, var_24
  loc_00DEF0BC: test eax, eax
  loc_00DEF0BE: jz 00DEF1F1h
  loc_00DEF0C4: cmp [eax], 0001h
  loc_00DEF0C8: jnz 00DEF1F1h
  loc_00DEF0CE: mov edi, var_80
  loc_00DEF0D1: mov edx, [eax+00000014h]
  loc_00DEF0D4: mov ecx, [eax+00000010h]
  loc_00DEF0D7: mov esi, ebx
  loc_00DEF0D9: sub esi, edi
  loc_00DEF0DB: jo 00DEF6D8h
  loc_00DEF0E1: sub esi, edx
  loc_00DEF0E3: cmp esi, ecx
  loc_00DEF0E5: jb 00DEF0EDh
  loc_00DEF0E7: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF0ED: shl esi, 02h
  loc_00DEF0F0: jmp 00DEF1F9h
  loc_00DEF0F5: mov ecx, var_38
  loc_00DEF0F8: xor esi, esi
  loc_00DEF0FA: cmp esi, ecx
  loc_00DEF0FC: jg 00DEF02Bh
  loc_00DEF102: mov ecx, var_24
  loc_00DEF105: test ecx, ecx
  loc_00DEF107: jz 00DEF13Eh
  loc_00DEF109: cmp [ecx], 0001h
  loc_00DEF10D: jnz 00DEF13Eh
  loc_00DEF10F: mov edx, [ecx+00000014h]
  loc_00DEF112: mov eax, esi
  loc_00DEF114: sub eax, edx
  loc_00DEF116: mov edx, [ecx+00000010h]
  loc_00DEF119: cmp eax, edx
  loc_00DEF11B: mov var_8C, eax
  loc_00DEF121: jb 00DEF12Fh
  loc_00DEF123: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF129: mov eax, var_8C
  loc_00DEF12F: lea edx, [eax*4]
  loc_00DEF136: mov var_F8, edx
  loc_00DEF13C: jmp 00DEF14Ah
  loc_00DEF13E: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF144: mov var_F8, eax
  loc_00DEF14A: mov eax, var_18
  loc_00DEF14D: mov ecx, var_38
  loc_00DEF150: imul eax, esi
  loc_00DEF153: jo 00DEF6D8h
  loc_00DEF159: cdq
  loc_00DEF15A: idiv ecx
  loc_00DEF15C: mov edx, eax
  loc_00DEF15E: mov eax, var_78
  loc_00DEF161: add edx, ebx
  loc_00DEF163: mov ebx, var_1C
  loc_00DEF166: jo 00DEF6D8h
  loc_00DEF16C: imul edx, edx, 00000100h
  loc_00DEF172: jo 00DEF6D8h
  loc_00DEF178: imul eax, esi
  loc_00DEF17B: jo 00DEF6D8h
  loc_00DEF181: mov var_FC, edx
  loc_00DEF187: cdq
  loc_00DEF188: idiv ecx
  loc_00DEF18A: add eax, ebx
  loc_00DEF18C: mov ebx, var_FC
  loc_00DEF192: jo 00DEF6D8h
  loc_00DEF198: add ebx, eax
  loc_00DEF19A: mov eax, var_7C
  loc_00DEF19D: jo 00DEF6D8h
  loc_00DEF1A3: imul eax, esi
  loc_00DEF1A6: jo 00DEF6D8h
  loc_00DEF1AC: cdq
  loc_00DEF1AD: idiv ecx
  loc_00DEF1AF: add eax, edi
  loc_00DEF1B1: jo 00DEF6D8h
  loc_00DEF1B7: imul eax, eax, 00010000h
  loc_00DEF1BD: jo 00DEF6D8h
  loc_00DEF1C3: add ebx, eax
  loc_00DEF1C5: mov eax, var_24
  loc_00DEF1C8: jo 00DEF6D8h
  loc_00DEF1CE: mov edx, [eax+0000000Ch]
  loc_00DEF1D1: mov eax, var_F8
  loc_00DEF1D7: mov [edx+eax], ebx
  loc_00DEF1DA: mov ebx, var_48
  loc_00DEF1DD: mov eax, 00000001h
  loc_00DEF1E2: add eax, esi
  loc_00DEF1E4: jo 00DEF6D8h
  loc_00DEF1EA: mov esi, eax
  loc_00DEF1EC: jmp 00DEF0FAh
  loc_00DEF1F1: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF1F7: mov esi, eax
  loc_00DEF1F9: mov ecx, var_44
  loc_00DEF1FC: test ecx, ecx
  loc_00DEF1FE: jz 00DEF226h
  loc_00DEF200: cmp [ecx], 0001h
  loc_00DEF204: jnz 00DEF226h
  loc_00DEF206: mov edx, [ecx+00000014h]
  loc_00DEF209: mov eax, [ecx+00000010h]
  loc_00DEF20C: mov edi, ebx
  loc_00DEF20E: sub edi, edx
  loc_00DEF210: cmp edi, eax
  loc_00DEF212: jb 00DEF21Dh
  loc_00DEF214: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF21A: mov ecx, var_44
  loc_00DEF21D: lea eax, [edi*4]
  loc_00DEF224: jmp 00DEF22Fh
  loc_00DEF226: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF22C: mov ecx, var_44
  loc_00DEF22F: mov edx, var_24
  loc_00DEF232: mov ecx, [ecx+0000000Ch]
  loc_00DEF235: mov edx, [edx+0000000Ch]
  loc_00DEF238: mov edx, [edx+esi]
  loc_00DEF23B: mov [ecx+eax], edx
  loc_00DEF23E: mov edx, var_80
  loc_00DEF241: mov ecx, var_B4
  loc_00DEF247: mov eax, 00000001h
  loc_00DEF24C: add eax, ebx
  loc_00DEF24E: jo 00DEF6D8h
  loc_00DEF254: mov ebx, eax
  loc_00DEF256: mov eax, var_28
  loc_00DEF259: jmp 00DEF0B1h
  loc_00DEF25E: mov edi, arg_14
  loc_00DEF261: mov esi, var_2C
  loc_00DEF264: add edx, edi
  loc_00DEF266: mov edi, var_38
  loc_00DEF269: mov ecx, 00000001h
  loc_00DEF26E: jo 00DEF6D8h
  loc_00DEF274: add ecx, esi
  loc_00DEF276: mov var_80, edx
  loc_00DEF279: jo 00DEF6D8h
  loc_00DEF27F: mov var_2C, ecx
  loc_00DEF282: jmp 00DEF097h
  loc_00DEF287: mov edx, var_80
  loc_00DEF28A: mov var_2C, eax
  loc_00DEF28D: xor ecx, ecx
  loc_00DEF28F: cmp eax, ecx
  loc_00DEF291: jl 00DEF5ACh
  loc_00DEF297: mov ecx, edx
  loc_00DEF299: mov ebx, edx
  loc_00DEF29B: add ecx, edi
  loc_00DEF29D: jo 00DEF6D8h
  loc_00DEF2A3: mov var_C4, ecx
  loc_00DEF2A9: cmp ebx, ecx
  loc_00DEF2AB: jg 00DEF348h
  loc_00DEF2B1: mov ecx, var_24
  loc_00DEF2B4: test ecx, ecx
  loc_00DEF2B6: jz 00DEF2DBh
  loc_00DEF2B8: cmp [ecx], 0001h
  loc_00DEF2BC: jnz 00DEF2DBh
  loc_00DEF2BE: mov edx, [ecx+00000014h]
  loc_00DEF2C1: mov esi, eax
  loc_00DEF2C3: mov eax, [ecx+00000010h]
  loc_00DEF2C6: sub esi, edx
  loc_00DEF2C8: cmp esi, eax
  loc_00DEF2CA: jb 00DEF2D2h
  loc_00DEF2CC: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF2D2: lea edi, [esi*4]
  loc_00DEF2D9: jmp 00DEF2E3h
  loc_00DEF2DB: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF2E1: mov edi, eax
  loc_00DEF2E3: mov ecx, var_44
  loc_00DEF2E6: test ecx, ecx
  loc_00DEF2E8: jz 00DEF310h
  loc_00DEF2EA: cmp [ecx], 0001h
  loc_00DEF2EE: jnz 00DEF310h
  loc_00DEF2F0: mov edx, [ecx+00000014h]
  loc_00DEF2F3: mov eax, [ecx+00000010h]
  loc_00DEF2F6: mov esi, ebx
  loc_00DEF2F8: sub esi, edx
  loc_00DEF2FA: cmp esi, eax
  loc_00DEF2FC: jb 00DEF307h
  loc_00DEF2FE: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF304: mov ecx, var_44
  loc_00DEF307: lea eax, [esi*4]
  loc_00DEF30E: jmp 00DEF319h
  loc_00DEF310: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF316: mov ecx, var_44
  loc_00DEF319: mov edx, var_24
  loc_00DEF31C: mov ecx, [ecx+0000000Ch]
  loc_00DEF31F: mov edx, [edx+0000000Ch]
  loc_00DEF322: mov edx, [edx+edi]
  loc_00DEF325: mov [ecx+eax], edx
  loc_00DEF328: mov edx, var_80
  loc_00DEF32B: mov ecx, var_C4
  loc_00DEF331: mov eax, 00000001h
  loc_00DEF336: add eax, ebx
  loc_00DEF338: jo 00DEF6D8h
  loc_00DEF33E: mov ebx, eax
  loc_00DEF340: mov eax, var_2C
  loc_00DEF343: jmp 00DEF2A9h
  loc_00DEF348: mov edi, arg_14
  loc_00DEF34B: add edx, edi
  loc_00DEF34D: mov edi, var_38
  loc_00DEF350: jo 00DEF6D8h
  loc_00DEF356: or ecx, FFFFFFFFh
  loc_00DEF359: mov var_80, edx
  loc_00DEF35C: add ecx, eax
  loc_00DEF35E: jo 00DEF6D8h
  loc_00DEF364: mov eax, ecx
  loc_00DEF366: mov var_2C, eax
  loc_00DEF369: jmp 00DEF28Dh
  loc_00DEF36E: mov ecx, eax
  loc_00DEF370: mov edi, var_40
  loc_00DEF373: imul ecx, ebx
  loc_00DEF376: jo 00DEF6D8h
  loc_00DEF37C: add eax, 00000001h
  loc_00DEF37F: mov var_2C, 00000001h
  loc_00DEF386: mov edx, var_2C
  loc_00DEF389: mov var_80, ecx
  loc_00DEF38C: jo 00DEF6D8h
  loc_00DEF392: mov var_CC, eax
  loc_00DEF398: cmp edx, var_CC
  loc_00DEF39E: jg 00DEF5ACh
  loc_00DEF3A4: mov esi, var_38
  loc_00DEF3A7: mov eax, ecx
  loc_00DEF3A9: add eax, esi
  loc_00DEF3AB: mov ebx, ecx
  loc_00DEF3AD: jo 00DEF6D8h
  loc_00DEF3B3: mov var_D4, eax
  loc_00DEF3B9: cmp ebx, eax
  loc_00DEF3BB: jg 00DEF46Bh
  loc_00DEF3C1: mov eax, var_24
  loc_00DEF3C4: test eax, eax
  loc_00DEF3C6: jz 00DEF3EBh
  loc_00DEF3C8: cmp [eax], 0001h
  loc_00DEF3CC: jnz 00DEF3EBh
  loc_00DEF3CE: mov edx, [eax+00000014h]
  loc_00DEF3D1: mov ecx, [eax+00000010h]
  loc_00DEF3D4: mov esi, edi
  loc_00DEF3D6: sub esi, edx
  loc_00DEF3D8: cmp esi, ecx
  loc_00DEF3DA: jb 00DEF3E2h
  loc_00DEF3DC: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF3E2: lea eax, [esi*4]
  loc_00DEF3E9: jmp 00DEF3F1h
  loc_00DEF3EB: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF3F1: mov ecx, var_44
  loc_00DEF3F4: mov var_100, eax
  loc_00DEF3FA: test ecx, ecx
  loc_00DEF3FC: jz 00DEF424h
  loc_00DEF3FE: cmp [ecx], 0001h
  loc_00DEF402: jnz 00DEF424h
  loc_00DEF404: mov edx, [ecx+00000014h]
  loc_00DEF407: mov eax, [ecx+00000010h]
  loc_00DEF40A: mov esi, ebx
  loc_00DEF40C: sub esi, edx
  loc_00DEF40E: cmp esi, eax
  loc_00DEF410: jb 00DEF41Bh
  loc_00DEF412: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF418: mov ecx, var_44
  loc_00DEF41B: lea eax, [esi*4]
  loc_00DEF422: jmp 00DEF42Dh
  loc_00DEF424: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF42A: mov ecx, var_44
  loc_00DEF42D: mov edx, var_24
  loc_00DEF430: mov esi, var_100
  loc_00DEF436: mov ecx, [ecx+0000000Ch]
  loc_00DEF439: add edi, 00000001h
  loc_00DEF43C: mov edx, [edx+0000000Ch]
  loc_00DEF43F: jo 00DEF6D8h
  loc_00DEF445: mov edx, [edx+esi]
  loc_00DEF448: mov [ecx+eax], edx
  loc_00DEF44B: mov ecx, var_80
  loc_00DEF44E: mov edx, var_2C
  loc_00DEF451: mov eax, 00000001h
  loc_00DEF456: add eax, ebx
  loc_00DEF458: jo 00DEF6D8h
  loc_00DEF45E: mov ebx, eax
  loc_00DEF460: mov eax, var_D4
  loc_00DEF466: jmp 00DEF3B9h
  loc_00DEF46B: mov edi, arg_14
  loc_00DEF46E: mov eax, 00000001h
  loc_00DEF473: sub ecx, edi
  loc_00DEF475: mov edi, edx
  loc_00DEF477: jo 00DEF6D8h
  loc_00DEF47D: add eax, edx
  loc_00DEF47F: mov var_80, ecx
  loc_00DEF482: jo 00DEF6D8h
  loc_00DEF488: mov edx, eax
  loc_00DEF48A: mov var_2C, edx
  loc_00DEF48D: jmp 00DEF398h
  loc_00DEF492: mov edi, var_40
  loc_00DEF495: add eax, 00000001h
  loc_00DEF498: jo 00DEF6D8h
  loc_00DEF49E: mov var_2C, 00000001h
  loc_00DEF4A5: mov var_DC, eax
  loc_00DEF4AB: mov eax, var_2C
  loc_00DEF4AE: mov var_80, 00000000h
  loc_00DEF4B5: cmp eax, var_DC
  loc_00DEF4BB: jg 00DEF5ACh
  loc_00DEF4C1: mov ebx, var_80
  loc_00DEF4C4: mov ecx, var_38
  loc_00DEF4C7: mov eax, ebx
  loc_00DEF4C9: add eax, ecx
  loc_00DEF4CB: jo 00DEF6D8h
  loc_00DEF4D1: mov var_E4, eax
  loc_00DEF4D7: cmp ebx, eax
  loc_00DEF4D9: jg 00DEF583h
  loc_00DEF4DF: mov eax, var_24
  loc_00DEF4E2: test eax, eax
  loc_00DEF4E4: jz 00DEF509h
  loc_00DEF4E6: cmp [eax], 0001h
  loc_00DEF4EA: jnz 00DEF509h
  loc_00DEF4EC: mov edx, [eax+00000014h]
  loc_00DEF4EF: mov ecx, [eax+00000010h]
  loc_00DEF4F2: mov esi, edi
  loc_00DEF4F4: sub esi, edx
  loc_00DEF4F6: cmp esi, ecx
  loc_00DEF4F8: jb 00DEF500h
  loc_00DEF4FA: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF500: lea eax, [esi*4]
  loc_00DEF507: jmp 00DEF50Fh
  loc_00DEF509: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF50F: mov ecx, var_44
  loc_00DEF512: mov var_104, eax
  loc_00DEF518: test ecx, ecx
  loc_00DEF51A: jz 00DEF542h
  loc_00DEF51C: cmp [ecx], 0001h
  loc_00DEF520: jnz 00DEF542h
  loc_00DEF522: mov edx, [ecx+00000014h]
  loc_00DEF525: mov eax, [ecx+00000010h]
  loc_00DEF528: mov esi, ebx
  loc_00DEF52A: sub esi, edx
  loc_00DEF52C: cmp esi, eax
  loc_00DEF52E: jb 00DEF539h
  loc_00DEF530: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF536: mov ecx, var_44
  loc_00DEF539: lea eax, [esi*4]
  loc_00DEF540: jmp 00DEF54Bh
  loc_00DEF542: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF548: mov ecx, var_44
  loc_00DEF54B: mov edx, var_24
  loc_00DEF54E: mov esi, var_104
  loc_00DEF554: mov ecx, [ecx+0000000Ch]
  loc_00DEF557: add edi, 00000001h
  loc_00DEF55A: mov edx, [edx+0000000Ch]
  loc_00DEF55D: jo 00DEF6D8h
  loc_00DEF563: mov edx, [edx+esi]
  loc_00DEF566: mov [ecx+eax], edx
  loc_00DEF569: mov eax, 00000001h
  loc_00DEF56E: add eax, ebx
  loc_00DEF570: jo 00DEF6D8h
  loc_00DEF576: mov ebx, eax
  loc_00DEF578: mov eax, var_E4
  loc_00DEF57E: jmp 00DEF4D7h
  loc_00DEF583: mov eax, var_80
  loc_00DEF586: mov edx, arg_14
  loc_00DEF589: mov edi, var_2C
  loc_00DEF58C: add eax, edx
  loc_00DEF58E: jo 00DEF6D8h
  loc_00DEF594: mov var_80, eax
  loc_00DEF597: mov eax, 00000001h
  loc_00DEF59C: add eax, edi
  loc_00DEF59E: jo 00DEF6D8h
  loc_00DEF5A4: mov var_2C, eax
  loc_00DEF5A7: jmp 00DEF4B5h
  loc_00DEF5AC: mov esi, Me
  loc_00DEF5AF: mov ebx, arg_14
  loc_00DEF5B2: mov edi, arg_18
  loc_00DEF5B5: lea edx, var_88
  loc_00DEF5BB: mov ecx, [esi]
  loc_00DEF5BD: push edx
  loc_00DEF5BE: push esi
  loc_00DEF5BF: mov var_70, 00000028h
  loc_00DEF5C6: mov var_64, 0001h
  loc_00DEF5CC: mov var_62, 0020h
  loc_00DEF5D2: mov var_6C, ebx
  loc_00DEF5D5: mov var_68, edi
  loc_00DEF5D8: call [ecx+000000D8h]
  loc_00DEF5DE: test eax, eax
  loc_00DEF5E0: fnclex
  loc_00DEF5E2: jge 00DEF5F6h
  loc_00DEF5E4: push 000000D8h
  loc_00DEF5E9: push 006DCB00h
  loc_00DEF5EE: push esi
  loc_00DEF5EF: push eax
  loc_00DEF5F0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEF5F6: mov eax, var_44
  loc_00DEF5F9: lea ecx, var_84
  loc_00DEF5FF: push eax
  loc_00DEF600: push ecx
  loc_00DEF601: call [004011DCh] ; __vbaAryLock
  loc_00DEF607: mov ecx, var_84
  loc_00DEF60D: test ecx, ecx
  loc_00DEF60F: jz 00DEF638h
  loc_00DEF611: cmp [ecx], 0001h
  loc_00DEF615: jnz 00DEF638h
  loc_00DEF617: mov esi, [ecx+00000014h]
  loc_00DEF61A: mov eax, [ecx+00000010h]
  loc_00DEF61D: neg esi
  loc_00DEF61F: cmp esi, eax
  loc_00DEF621: jb 00DEF62Fh
  loc_00DEF623: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF629: mov ecx, var_84
  loc_00DEF62F: lea eax, [esi*4]
  loc_00DEF636: jmp 00DEF644h
  loc_00DEF638: call [00401100h] ; __vbaGenerateBoundsError
  loc_00DEF63E: mov ecx, var_84
  loc_00DEF644: mov ecx, [ecx+0000000Ch]
  loc_00DEF647: push 00CC0020h
  loc_00DEF64C: lea edx, var_70
  loc_00DEF64F: push 00000000h
  loc_00DEF651: add ecx, eax
  loc_00DEF653: mov eax, arg_C
  loc_00DEF656: push edx
  loc_00DEF657: mov edx, arg_10
  loc_00DEF65A: push ecx
  loc_00DEF65B: mov ecx, var_88
  loc_00DEF661: push edi
  loc_00DEF662: push ebx
  loc_00DEF663: push 00000000h
  loc_00DEF665: push 00000000h
  loc_00DEF667: push edi
  loc_00DEF668: push ebx
  loc_00DEF669: push edx
  loc_00DEF66A: push eax
  loc_00DEF66B: push ecx
  loc_00DEF66C: call 006DD388h ; StretchDIBits(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v, %x10v, %x11v, %x12v, %x13v)
  loc_00DEF671: call [00401074h] ; __vbaSetSystemError
  loc_00DEF677: lea edx, var_84
  loc_00DEF67D: push edx
  loc_00DEF67E: call [00401244h] ; __vbaAryUnlock
  loc_00DEF684: push 00DEF6B0h
  loc_00DEF689: jmp 00DEF699h
  loc_00DEF68B: lea eax, var_84
  loc_00DEF691: push eax
  loc_00DEF692: call [00401244h] ; __vbaAryUnlock
  loc_00DEF698: ret
  loc_00DEF699: mov esi, [00401090h] ; __vbaAryDestruct
  loc_00DEF69F: lea ecx, var_24
  loc_00DEF6A2: push ecx
  loc_00DEF6A3: push 00000000h
  loc_00DEF6A5: call __vbaAryDestruct
  loc_00DEF6A7: lea edx, var_44
  loc_00DEF6AA: push edx
  loc_00DEF6AB: push 00000000h
  loc_00DEF6AD: call __vbaAryDestruct
  loc_00DEF6AF: ret
  loc_00DEF6B0: mov ecx, var_10
  loc_00DEF6B3: pop edi
  loc_00DEF6B4: pop esi
  loc_00DEF6B5: xor eax, eax
  loc_00DEF6B7: mov fs:[00000000h], ecx
  loc_00DEF6BE: pop ebx
  loc_00DEF6BF: mov esp, ebp
  loc_00DEF6C1: pop ebp
  loc_00DEF6C2: retn 0020h
End Function

Private Sub Proc_2_110_DEF6E0(arg_C) 'DEF6E0
  loc_00DEF6E0: push ecx
  loc_00DEF6E1: mov ecx, Me
  loc_00DEF6E5: lea eax, [esp]
  loc_00DEF6E9: push esi
  loc_00DEF6EA: push eax
  loc_00DEF6EB: mov edx, [ecx]
  loc_00DEF6ED: mov eax, arg_C
  loc_00DEF6F1: push edx
  loc_00DEF6F2: push eax
  loc_00DEF6F3: mov Me, 00000000h
  loc_00DEF6FB: call 006DD7E0h ; OleTranslateColor()
  loc_00DEF700: mov esi, eax
  loc_00DEF702: call [00401074h] ; __vbaSetSystemError
  loc_00DEF708: test esi, esi
  loc_00DEF70A: pop esi
  loc_00DEF70B: jz 00DEF71Ch
  loc_00DEF70D: mov ecx, arg_C
  loc_00DEF711: or eax, FFFFFFFFh
  loc_00DEF714: mov [ecx], eax
  loc_00DEF716: xor eax, eax
  loc_00DEF718: pop ecx
  loc_00DEF719: retn 0010h
End Sub

Private Sub Proc_2_111_DEF730
  loc_00DEF730: sub esp, 00000008h
  loc_00DEF733: push esi
  loc_00DEF734: mov esi, Me
  loc_00DEF738: mov [esp+00000008h], 00000000h
  loc_00DEF740: mov var_4, 00000000h
  loc_00DEF748: mov eax, [esi+00000010h]
  loc_00DEF74B: push edi
  loc_00DEF74C: push eax
  loc_00DEF74D: mov ecx, [eax]
  loc_00DEF74F: call [ecx+0000035Ch]
  loc_00DEF755: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEF75B: test eax, eax
  loc_00DEF75D: fnclex
  loc_00DEF75F: jge 00DEF772h
  loc_00DEF761: mov edx, [esi+00000010h]
  loc_00DEF764: push 0000035Ch
  loc_00DEF769: push 006DCB00h
  loc_00DEF76E: push edx
  loc_00DEF76F: push eax
  loc_00DEF770: call edi
  loc_00DEF772: mov eax, [esi]
  loc_00DEF774: lea ecx, [esp+00000008h]
  loc_00DEF778: push ecx
  loc_00DEF779: push esi
  loc_00DEF77A: call [eax+00000108h]
  loc_00DEF780: test eax, eax
  loc_00DEF782: fnclex
  loc_00DEF784: jge 00DEF794h
  loc_00DEF786: push 00000108h
  loc_00DEF78B: push 006DCB00h
  loc_00DEF790: push esi
  loc_00DEF791: push eax
  loc_00DEF792: call edi
  loc_00DEF794: fld real4 ptr [esp+00000008h]
  loc_00DEF798: push ebx
  loc_00DEF799: mov ebx, [00401208h] ; __vbaFpI4
  loc_00DEF79F: call ebx
  loc_00DEF7A1: mov edx, [esi]
  loc_00DEF7A3: mov [esi+000001D0h], eax
  loc_00DEF7A9: lea eax, Proc_2_111_DEF730
  loc_00DEF7AD: push eax
  loc_00DEF7AE: push esi
  loc_00DEF7AF: call [edx+00000100h]
  loc_00DEF7B5: test eax, eax
  loc_00DEF7B7: fnclex
  loc_00DEF7B9: jge 00DEF7C9h
  loc_00DEF7BB: push 00000100h
  loc_00DEF7C0: push 006DCB00h
  loc_00DEF7C5: push esi
  loc_00DEF7C6: push eax
  loc_00DEF7C7: call edi
  loc_00DEF7C9: fld real4 ptr Proc_2_111_DEF730
  loc_00DEF7CD: call ebx
  loc_00DEF7CF: mov ecx, [esi+000001D0h]
  loc_00DEF7D5: lea edx, [esi+000000ACh]
  loc_00DEF7DB: push ecx
  loc_00DEF7DC: push eax
  loc_00DEF7DD: push 00000000h
  loc_00DEF7DF: push 00000000h
  loc_00DEF7E1: push edx
  loc_00DEF7E2: mov [esi+000001D4h], eax
  loc_00DEF7E8: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DEF7ED: call [00401074h] ; __vbaSetSystemError
  loc_00DEF7F3: mov eax, [esi+00000064h]
  loc_00DEF7F6: mov edi, 00000002h
  loc_00DEF7FB: test eax, eax
  loc_00DEF7FD: pop ebx
  loc_00DEF7FE: jz 00DEF816h
  loc_00DEF800: mov eax, [esi+00000044h]
  loc_00DEF803: test eax, eax
  loc_00DEF805: jz 00DEF816h
  loc_00DEF807: cmp eax, 00000006h
  loc_00DEF80A: jz 00DEF816h
  loc_00DEF80C: cmp [esi+0000006Ch], 0000h
  loc_00DEF811: jz 00DEF816h
  loc_00DEF813: mov [esi+00000048h], edi
  loc_00DEF816: mov eax, [esi+00000044h]
  loc_00DEF819: cmp eax, 0000000Dh
  loc_00DEF81C: ja 00DEF98Dh
  loc_00DEF822: jmp [eax*4+00DEF998h]
  loc_00DEF829: mov ecx, [esi+00000048h]
  loc_00DEF82C: mov eax, [esi]
  loc_00DEF82E: push ecx
  loc_00DEF82F: push esi
  loc_00DEF830: call [eax+00000944h]
  loc_00DEF836: pop edi
  loc_00DEF837: xor eax, eax
  loc_00DEF839: pop esi
  loc_00DEF83A: add esp, 00000008h
  loc_00DEF83D: retn 0004h
End Sub

Private Sub Proc_2_112_DEF9D0(arg_C, arg_10) 'DEF9D0
  loc_00DEF9D0: sub esp, 00000008h
  loc_00DEF9D3: push ebx
  loc_00DEF9D4: push ebp
  loc_00DEF9D5: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00DEF9DB: push esi
  loc_00DEF9DC: mov esi, arg_10
  loc_00DEF9E0: push edi
  loc_00DEF9E1: xor edi, edi
  loc_00DEF9E3: mov eax, [esi+000000A4h]
  loc_00DEF9E9: mov Me, edi
  loc_00DEF9ED: cmp eax, edi
  loc_00DEF9EF: mov arg_C, edi
  loc_00DEF9F3: jz 00DEF9FDh
  loc_00DEF9F5: push eax
  loc_00DEF9F6: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEF9FB: call ebx
  loc_00DEF9FD: mov eax, [esi+00000044h]
  loc_00DEFA00: mov ebp, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEFA06: add eax, FFFFFFFEh
  loc_00DEFA09: cmp eax, 00000008h
  loc_00DEFA0C: ja 00DEFA96h
  loc_00DEFA12: jmp [eax*4+00DEFB60h]
  loc_00DEFA19: mov eax, [esi+000001D0h]
  loc_00DEFA1F: mov ecx, [esi+000001D4h]
  loc_00DEFA25: add eax, 00000001h
  loc_00DEFA28: push 00000003h
  loc_00DEFA2A: jo 00DEFB84h
  loc_00DEFA30: add ecx, 00000001h
  loc_00DEFA33: push 00000003h
  loc_00DEFA35: push eax
  loc_00DEFA36: jo 00DEFB84h
  loc_00DEFA3C: push ecx
  loc_00DEFA3D: push edi
  loc_00DEFA3E: push edi
  loc_00DEFA3F: call 006DD604h ; CreateRoundRectRgn(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v)
  loc_00DEFA44: mov Me, eax
  loc_00DEFA48: call ebx
  loc_00DEFA4A: mov edx, Me
  loc_00DEFA4E: mov [esi+000000A4h], edx
  loc_00DEFA54: jmp 00DEFB0Dh
  loc_00DEFA59: mov eax, [esi+000001D0h]
  loc_00DEFA5F: mov ecx, [esi+000001D4h]
  loc_00DEFA65: add eax, 00000001h
  loc_00DEFA68: push 00000004h
  loc_00DEFA6A: jo 00DEFB84h
  loc_00DEFA70: add ecx, 00000001h
  loc_00DEFA73: push 00000004h
  loc_00DEFA75: push eax
  loc_00DEFA76: jo 00DEFB84h
  loc_00DEFA7C: push ecx
  loc_00DEFA7D: push edi
  loc_00DEFA7E: push edi
  loc_00DEFA7F: call 006DD604h ; CreateRoundRectRgn(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v)
  loc_00DEFA84: mov Me, eax
  loc_00DEFA88: call ebx
  loc_00DEFA8A: mov edx, Me
  loc_00DEFA8E: mov [esi+000000A4h], edx
  loc_00DEFA94: jmp 00DEFB0Dh
  loc_00DEFA96: mov eax, [esi+00000010h]
  loc_00DEFA99: lea edx, Me
  loc_00DEFA9D: push edx
  loc_00DEFA9E: push eax
  loc_00DEFA9F: mov ecx, [eax]
  loc_00DEFAA1: call [ecx+00000100h]
  loc_00DEFAA7: cmp eax, edi
  loc_00DEFAA9: fnclex
  loc_00DEFAAB: jge 00DEFABEh
  loc_00DEFAAD: mov ecx, [esi+00000010h]
  loc_00DEFAB0: push 00000100h
  loc_00DEFAB5: push 006DCB00h
  loc_00DEFABA: push ecx
  loc_00DEFABB: push eax
  loc_00DEFABC: call ebp
  loc_00DEFABE: mov eax, [esi+00000010h]
  loc_00DEFAC1: lea ecx, arg_C
  loc_00DEFAC5: push ecx
  loc_00DEFAC6: push eax
  loc_00DEFAC7: mov edx, [eax]
  loc_00DEFAC9: call [edx+00000108h]
  loc_00DEFACF: cmp eax, edi
  loc_00DEFAD1: fnclex
  loc_00DEFAD3: jge 00DEFAE6h
  loc_00DEFAD5: mov edx, [esi+00000010h]
  loc_00DEFAD8: push 00000108h
  loc_00DEFADD: push 006DCB00h
  loc_00DEFAE2: push edx
  loc_00DEFAE3: push eax
  loc_00DEFAE4: call ebp
  loc_00DEFAE6: fld real4 ptr arg_C
  loc_00DEFAEA: mov edi, [00401208h] ; __vbaFpI4
  loc_00DEFAF0: call edi
  loc_00DEFAF2: fld real4 ptr Me
  loc_00DEFAF6: push eax
  loc_00DEFAF7: call edi
  loc_00DEFAF9: push eax
  loc_00DEFAFA: push 00000000h
  loc_00DEFAFC: push 00000000h
  loc_00DEFAFE: call 006DD6D4h ; CreateRectRgn(%x1v, %x2v, %x3v, %x4v)
  loc_00DEFB03: mov edi, eax
  loc_00DEFB05: call ebx
  loc_00DEFB07: mov [esi+000000A4h], edi
  loc_00DEFB0D: mov eax, [esi+00000010h]
  loc_00DEFB10: lea edx, Me
  loc_00DEFB14: push edx
  loc_00DEFB15: push eax
  loc_00DEFB16: mov ecx, [eax]
  loc_00DEFB18: call [ecx+00000058h]
  loc_00DEFB1B: test eax, eax
  loc_00DEFB1D: fnclex
  loc_00DEFB1F: jge 00DEFB2Fh
  loc_00DEFB21: mov ecx, [esi+00000010h]
  loc_00DEFB24: push 00000058h
  loc_00DEFB26: push 006DCB00h
  loc_00DEFB2B: push ecx
  loc_00DEFB2C: push eax
  loc_00DEFB2D: call ebp
  loc_00DEFB2F: mov edx, [esi+000000A4h]
  loc_00DEFB35: mov eax, Me
  loc_00DEFB39: push FFFFFFFFh
  loc_00DEFB3B: push edx
  loc_00DEFB3C: push eax
  loc_00DEFB3D: call 006DDCD8h ; SetWindowRgn(%x1v, %x2v, %x3v)
  loc_00DEFB42: call ebx
  loc_00DEFB44: mov ecx, [esi+000000A4h]
  loc_00DEFB4A: push ecx
  loc_00DEFB4B: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEFB50: call ebx
  loc_00DEFB52: pop edi
  loc_00DEFB53: pop esi
  loc_00DEFB54: pop ebp
  loc_00DEFB55: xor eax, eax
  loc_00DEFB57: pop ebx
  loc_00DEFB58: add esp, 00000008h
  loc_00DEFB5B: retn 0004h
End Sub

Private Sub Proc_2_113_DEFB90(arg_C) 'DEFB90
  loc_00DEFB90: push ebp
  loc_00DEFB91: mov ebp, esp
  loc_00DEFB93: sub esp, 00000008h
  loc_00DEFB96: push 00402836h ; __vbaExceptHandler
  loc_00DEFB9B: mov eax, fs:[00000000h]
  loc_00DEFBA1: push eax
  loc_00DEFBA2: mov fs:[00000000h], esp
  loc_00DEFBA9: sub esp, 00000024h
  loc_00DEFBAC: push ebx
  loc_00DEFBAD: push esi
  loc_00DEFBAE: push edi
  loc_00DEFBAF: mov var_8, esp
  loc_00DEFBB2: mov var_4, 00401390h
  loc_00DEFBB9: mov esi, Me
  loc_00DEFBBC: lea ecx, var_2C
  loc_00DEFBBF: lea edx, var_28
  loc_00DEFBC2: push ecx
  loc_00DEFBC3: mov eax, [esi]
  loc_00DEFBC5: xor edi, edi
  loc_00DEFBC7: push edx
  loc_00DEFBC8: push esi
  loc_00DEFBC9: mov var_1C, edi
  loc_00DEFBCC: mov var_24, edi
  loc_00DEFBCF: mov var_2C, edi
  loc_00DEFBD2: mov var_28, 0000000Eh
  loc_00DEFBD9: call [eax+0000091Ch]
  loc_00DEFBDF: mov eax, [esi]
  loc_00DEFBE1: mov ebx, var_2C
  loc_00DEFBE4: lea ecx, var_28
  loc_00DEFBE7: push ecx
  loc_00DEFBE8: push esi
  loc_00DEFBE9: call [eax+000000D8h]
  loc_00DEFBEF: cmp eax, edi
  loc_00DEFBF1: fnclex
  loc_00DEFBF3: jge 00DEFC07h
  loc_00DEFBF5: push 000000D8h
  loc_00DEFBFA: push 006DCB00h
  loc_00DEFBFF: push esi
  loc_00DEFC00: push eax
  loc_00DEFC01: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEFC07: mov edx, var_28
  loc_00DEFC0A: push ebx
  loc_00DEFC0B: push edx
  loc_00DEFC0C: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DEFC11: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DEFC17: mov var_2C, eax
  loc_00DEFC1A: call edi
  loc_00DEFC1C: mov eax, arg_C
  loc_00DEFC1F: push eax
  loc_00DEFC20: call [00401018h] ; __vbaStrI4
  loc_00DEFC26: mov edx, eax
  loc_00DEFC28: lea ecx, var_1C
  loc_00DEFC2B: call [00401228h] ; __vbaStrMove
  loc_00DEFC31: mov ecx, [esi]
  loc_00DEFC33: lea edx, var_28
  loc_00DEFC36: push edx
  loc_00DEFC37: push esi
  loc_00DEFC38: call [ecx+000000D8h]
  loc_00DEFC3E: test eax, eax
  loc_00DEFC40: fnclex
  loc_00DEFC42: jge 00DEFC56h
  loc_00DEFC44: push 000000D8h
  loc_00DEFC49: push 006DCB00h
  loc_00DEFC4E: push esi
  loc_00DEFC4F: push eax
  loc_00DEFC50: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEFC56: mov eax, var_1C
  loc_00DEFC59: add esi, 00000128h
  loc_00DEFC5F: push 00000010h
  loc_00DEFC61: push esi
  loc_00DEFC62: push 00000001h
  loc_00DEFC64: lea ecx, var_24
  loc_00DEFC67: push eax
  loc_00DEFC68: push ecx
  loc_00DEFC69: call [004011ECh] ; __vbaStrToAnsi
  loc_00DEFC6F: mov edx, var_28
  loc_00DEFC72: push eax
  loc_00DEFC73: push edx
  loc_00DEFC74: call 006DE2BCh ; DrawText(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DEFC79: call edi
  loc_00DEFC7B: mov eax, var_24
  loc_00DEFC7E: lea ecx, var_1C
  loc_00DEFC81: push eax
  loc_00DEFC82: push ecx
  loc_00DEFC83: call [00401160h] ; __vbaStrToUnicode
  loc_00DEFC89: lea ecx, var_24
  loc_00DEFC8C: call [00401258h] ; __vbaFreeStr
  loc_00DEFC92: push ebx
  loc_00DEFC93: call 006DD498h ; DeleteObject(%x1v)
  loc_00DEFC98: call edi
  loc_00DEFC9A: push 00DEFCB5h
  loc_00DEFC9F: jmp 00DEFCABh
  loc_00DEFCA1: lea ecx, var_24
  loc_00DEFCA4: call [00401258h] ; __vbaFreeStr
  loc_00DEFCAA: ret
  loc_00DEFCAB: lea ecx, var_1C
  loc_00DEFCAE: call [00401258h] ; __vbaFreeStr
  loc_00DEFCB4: ret
  loc_00DEFCB5: mov ecx, var_10
  loc_00DEFCB8: pop edi
  loc_00DEFCB9: pop esi
  loc_00DEFCBA: xor eax, eax
  loc_00DEFCBC: mov fs:[00000000h], ecx
  loc_00DEFCC3: pop ebx
  loc_00DEFCC4: mov esp, ebp
  loc_00DEFCC6: pop ebp
  loc_00DEFCC7: retn 0008h
End Sub

Private Sub Proc_2_114_DEFCD0(arg_C, arg_10) 'DEFCD0
  loc_00DEFCD0: push ebp
  loc_00DEFCD1: mov ebp, esp
  loc_00DEFCD3: sub esp, 00000008h
  loc_00DEFCD6: push 00402836h ; __vbaExceptHandler
  loc_00DEFCDB: mov eax, fs:[00000000h]
  loc_00DEFCE1: push eax
  loc_00DEFCE2: mov fs:[00000000h], esp
  loc_00DEFCE9: sub esp, 000000ACh
  loc_00DEFCEF: push ebx
  loc_00DEFCF0: push esi
  loc_00DEFCF1: push edi
  loc_00DEFCF2: mov var_8, esp
  loc_00DEFCF5: mov var_4, 004013A0h
  loc_00DEFCFC: mov ecx, 00000017h
  loc_00DEFD01: xor eax, eax
  loc_00DEFD03: lea edi, var_70
  loc_00DEFD06: push 006DE9E4h ; "Marlett"
  loc_00DEFD0B: repz stosd
  loc_00DEFD0D: mov ecx, 0000000Fh
  loc_00DEFD12: lea edi, var_B8
  loc_00DEFD18: repz stosd
  loc_00DEFD1A: push 006DE9F8h
  loc_00DEFD1F: mov var_74, eax
  loc_00DEFD22: call [0040106Ch] ; __vbaStrCat
  loc_00DEFD28: mov edx, eax
  loc_00DEFD2A: lea ecx, var_74
  loc_00DEFD2D: call [00401228h] ; __vbaStrMove
  loc_00DEFD33: push eax
  loc_00DEFD34: lea eax, var_54
  loc_00DEFD37: push eax
  loc_00DEFD38: push 00000020h
  loc_00DEFD3A: call [00401070h] ; __vbaLsetFixstr
  loc_00DEFD40: lea ecx, var_74
  loc_00DEFD43: call [00401258h] ; __vbaFreeStr
  loc_00DEFD49: mov ecx, arg_C
  loc_00DEFD4C: mov edx, [ecx]
  loc_00DEFD4E: mov ecx, 00000002h
  loc_00DEFD53: mov var_70, edx
  loc_00DEFD56: call [00401144h] ; __vbaUI1I2
  loc_00DEFD5C: mov var_59, al
  loc_00DEFD5F: lea eax, var_70
  loc_00DEFD62: lea ecx, var_B8
  loc_00DEFD68: push eax
  loc_00DEFD69: push ecx
  loc_00DEFD6A: push 006DD1BCh ; UDT_8_006DD1BC
  loc_00DEFD6F: call [0040113Ch] ; __vbaRecUniToAnsi
  loc_00DEFD75: push eax
  loc_00DEFD76: call 006DD8B4h ; CreateFontIndirect(%x1v)
  loc_00DEFD7B: mov esi, eax
  loc_00DEFD7D: call [00401074h] ; __vbaSetSystemError
  loc_00DEFD83: lea edx, var_B8
  loc_00DEFD89: lea eax, var_70
  loc_00DEFD8C: push edx
  loc_00DEFD8D: push eax
  loc_00DEFD8E: push 006DD1BCh ; UDT_8_006DD1BC
  loc_00DEFD93: call [00401058h] ; __vbaRecAnsiToUni
  loc_00DEFD99: mov var_14, esi
  loc_00DEFD9C: push 00DEFDAEh
  loc_00DEFDA1: jmp 00DEFDADh
  loc_00DEFDA3: lea ecx, var_74
  loc_00DEFDA6: call [00401258h] ; __vbaFreeStr
  loc_00DEFDAC: ret
  loc_00DEFDAD: ret
  loc_00DEFDAE: mov ecx, arg_10
  loc_00DEFDB1: mov edx, var_14
  loc_00DEFDB4: pop edi
  loc_00DEFDB5: pop esi
  loc_00DEFDB6: mov [ecx], edx
  loc_00DEFDB8: mov ecx, var_10
  loc_00DEFDBB: xor eax, eax
  loc_00DEFDBD: mov fs:[00000000h], ecx
  loc_00DEFDC4: pop ebx
  loc_00DEFDC5: mov esp, ebp
  loc_00DEFDC7: pop ebp
  loc_00DEFDC8: retn 000Ch
End Sub

Private Sub Proc_2_115_DEFDD0
  loc_00DEFDD0: push ebp
  loc_00DEFDD1: mov ebp, esp
  loc_00DEFDD3: sub esp, 00000008h
  loc_00DEFDD6: push 00402836h ; __vbaExceptHandler
  loc_00DEFDDB: mov eax, fs:[00000000h]
  loc_00DEFDE1: push eax
  loc_00DEFDE2: mov fs:[00000000h], esp
  loc_00DEFDE9: sub esp, 0000010Ch
  loc_00DEFDEF: push ebx
  loc_00DEFDF0: push esi
  loc_00DEFDF1: push edi
  loc_00DEFDF2: mov var_8, esp
  loc_00DEFDF5: mov var_4, 004013D8h
  loc_00DEFDFC: xor eax, eax
  loc_00DEFDFE: mov esi, Me
  loc_00DEFE01: mov var_24, eax
  loc_00DEFE04: xor ecx, ecx
  loc_00DEFE06: mov edx, [esi]
  loc_00DEFE08: mov var_20, eax
  loc_00DEFE0B: mov var_1C, eax
  loc_00DEFE0E: mov var_34, ecx
  loc_00DEFE11: mov var_18, eax
  loc_00DEFE14: mov var_30, ecx
  loc_00DEFE17: lea eax, var_C0
  loc_00DEFE1D: mov var_2C, ecx
  loc_00DEFE20: xor ebx, ebx
  loc_00DEFE22: push eax
  loc_00DEFE23: push esi
  loc_00DEFE24: mov var_28, ecx
  loc_00DEFE27: mov var_38, ebx
  loc_00DEFE2A: mov var_3C, ebx
  loc_00DEFE2D: mov var_4C, ebx
  loc_00DEFE30: mov var_5C, ebx
  loc_00DEFE33: mov var_6C, ebx
  loc_00DEFE36: mov var_7C, ebx
  loc_00DEFE39: mov var_8C, ebx
  loc_00DEFE3F: mov var_BC, ebx
  loc_00DEFE45: mov var_C0, ebx
  loc_00DEFE4B: mov var_C4, ebx
  loc_00DEFE51: mov var_C8, ebx
  loc_00DEFE57: mov var_CC, ebx
  loc_00DEFE5D: mov var_D0, ebx
  loc_00DEFE63: call [edx+00000100h]
  loc_00DEFE69: cmp eax, ebx
  loc_00DEFE6B: fnclex
  loc_00DEFE6D: jge 00DEFE81h
  loc_00DEFE6F: push 00000100h
  loc_00DEFE74: push 006DCB00h
  loc_00DEFE79: push esi
  loc_00DEFE7A: push eax
  loc_00DEFE7B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEFE81: fld real4 ptr var_C0
  loc_00DEFE87: mov edi, [00401208h] ; __vbaFpI4
  loc_00DEFE8D: call edi
  loc_00DEFE8F: mov ecx, [esi]
  loc_00DEFE91: lea edx, var_C0
  loc_00DEFE97: push edx
  loc_00DEFE98: push esi
  loc_00DEFE99: mov [esi+000001D4h], eax
  loc_00DEFE9F: call [ecx+00000108h]
  loc_00DEFEA5: cmp eax, ebx
  loc_00DEFEA7: fnclex
  loc_00DEFEA9: jge 00DEFEBDh
  loc_00DEFEAB: push 00000108h
  loc_00DEFEB0: push 006DCB00h
  loc_00DEFEB5: push esi
  loc_00DEFEB6: push eax
  loc_00DEFEB7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEFEBD: fld real4 ptr var_C0
  loc_00DEFEC3: call edi
  loc_00DEFEC5: mov [esi+000001D0h], eax
  loc_00DEFECB: mov eax, [esi+00000048h]
  loc_00DEFECE: cmp eax, 00000002h
  loc_00DEFED1: jz 00DEFF03h
  loc_00DEFED3: cmp [esi+00000064h], ebx
  loc_00DEFED6: jz 00DEFEDFh
  loc_00DEFED8: cmp [esi+0000006Ch], FFFFFFh
  loc_00DEFEDD: jz 00DEFF03h
  loc_00DEFEDF: cmp eax, 00000001h
  loc_00DEFEE2: jnz 00DEFEF4h
  loc_00DEFEE4: mov eax, [esi+00000150h]
  loc_00DEFEEA: push ebx
  loc_00DEFEEB: push eax
  loc_00DEFEEC: call [0040114Ch] ; __vbaObjIs
  loc_00DEFEF2: jmp 00DEFF30h
  loc_00DEFEF4: mov eax, [esi+0000014Ch]
  loc_00DEFEFA: lea edi, [esi+000001B4h]
  loc_00DEFF00: push eax
  loc_00DEFF01: jmp 00DEFF4Bh
  loc_00DEFF03: mov ecx, [esi+00000154h]
  loc_00DEFF09: mov edi, [0040114Ch] ; __vbaObjIs
  loc_00DEFF0F: push ebx
  loc_00DEFF10: push ecx
  loc_00DEFF11: call edi
  loc_00DEFF13: test ax, ax
  loc_00DEFF16: jnz 00DEFF26h
  loc_00DEFF18: mov edx, [esi+00000154h]
  loc_00DEFF1E: lea edi, [esi+000001B4h]
  loc_00DEFF24: jmp 00DEFF4Ah
  loc_00DEFF26: mov eax, [esi+00000150h]
  loc_00DEFF2C: push ebx
  loc_00DEFF2D: push eax
  loc_00DEFF2E: call edi
  loc_00DEFF30: test ax, ax
  loc_00DEFF33: lea edi, [esi+000001B4h]
  loc_00DEFF39: jnz 00DEFF44h
  loc_00DEFF3B: mov ecx, [esi+00000150h]
  loc_00DEFF41: push ecx
  loc_00DEFF42: jmp 00DEFF4Bh
  loc_00DEFF44: mov edx, [esi+0000014Ch]
  loc_00DEFF4A: push edx
  loc_00DEFF4B: push edi
  loc_00DEFF4C: call [004010B4h] ; __vbaObjSetAddref
  loc_00DEFF52: cmp [edi], ebx
  loc_00DEFF54: jnz 00DEFF62h
  loc_00DEFF56: push edi
  loc_00DEFF57: push 006DEA0Ch
  loc_00DEFF5C: call [004011A0h] ; __vbaNew2
  loc_00DEFF62: mov eax, [edi]
  loc_00DEFF64: push ebx
  loc_00DEFF65: push 00000005h
  loc_00DEFF67: lea ecx, var_4C
  loc_00DEFF6A: push eax
  loc_00DEFF6B: push ecx
  loc_00DEFF6C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DEFF72: add esp, 00000010h
  loc_00DEFF75: lea edx, var_C0
  loc_00DEFF7B: mov ecx, 00000003h
  loc_00DEFF80: mov var_84, 00000008h
  loc_00DEFF8A: push edx
  loc_00DEFF8B: mov eax, ecx
  loc_00DEFF8D: sub esp, 00000010h
  loc_00DEFF90: mov var_8C, ecx
  loc_00DEFF96: mov edx, esp
  loc_00DEFF98: sub esp, 00000010h
  loc_00DEFF9B: mov ebx, [esi]
  loc_00DEFF9D: mov [edx], ecx
  loc_00DEFF9F: mov ecx, var_98
  loc_00DEFFA5: mov [edx+00000004h], ecx
  loc_00DEFFA8: mov ecx, esp
  loc_00DEFFAA: mov [edx+00000008h], eax
  loc_00DEFFAD: mov eax, var_90
  loc_00DEFFB3: mov [edx+0000000Ch], eax
  loc_00DEFFB6: mov edx, var_8C
  loc_00DEFFBC: mov eax, var_88
  loc_00DEFFC2: mov [ecx], edx
  loc_00DEFFC4: mov edx, var_84
  loc_00DEFFCA: mov [ecx+00000004h], eax
  loc_00DEFFCD: mov eax, var_80
  loc_00DEFFD0: mov [ecx+00000008h], edx
  loc_00DEFFD3: mov [ecx+0000000Ch], eax
  loc_00DEFFD6: lea ecx, var_4C
  loc_00DEFFD9: push ecx
  loc_00DEFFDA: call [004011D0h] ; __vbaI4Var
  loc_00DEFFE0: mov var_E8, eax
  loc_00DEFFE6: fild real4 ptr var_E8
  loc_00DEFFEC: fstp real4 ptr var_EC
  loc_00DEFFF2: mov edx, var_EC
  loc_00DEFFF8: push edx
  loc_00DEFFF9: push esi
  loc_00DEFFFA: call [ebx+00000374h]
  loc_00DF0000: test eax, eax
  loc_00DF0002: fnclex
  loc_00DF0004: jge 00DF0018h
  loc_00DF0006: push 00000374h
  loc_00DF000B: push 006DCB00h
  loc_00DF0010: push esi
  loc_00DF0011: push eax
  loc_00DF0012: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF0018: fld real4 ptr var_C0
  loc_00DF001E: call [00401208h] ; __vbaFpI4
  loc_00DF0024: lea ecx, var_4C
  loc_00DF0027: mov [esi+00000174h], eax
  loc_00DF002D: call [00401024h] ; __vbaFreeVar
  loc_00DF0033: cmp [edi], 00000000h
  loc_00DF0036: jnz 00DF0044h
  loc_00DF0038: push edi
  loc_00DF0039: push 006DEA0Ch
  loc_00DF003E: call [004011A0h] ; __vbaNew2
  loc_00DF0044: mov eax, [edi]
  loc_00DF0046: push 00000000h
  loc_00DF0048: push 00000004h
  loc_00DF004A: lea ecx, var_4C
  loc_00DF004D: push eax
  loc_00DF004E: push ecx
  loc_00DF004F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DF0055: add esp, 00000010h
  loc_00DF0058: lea ebx, var_C0
  loc_00DF005E: mov ecx, 00000003h
  loc_00DF0063: mov edi, [esi]
  loc_00DF0065: push ebx
  loc_00DF0066: mov eax, ecx
  loc_00DF0068: sub esp, 00000010h
  loc_00DF006B: mov var_8C, ecx
  loc_00DF0071: mov ebx, esp
  loc_00DF0073: sub esp, 00000010h
  loc_00DF0076: mov edx, 00000008h
  loc_00DF007B: mov [ebx], ecx
  loc_00DF007D: mov ecx, var_98
  loc_00DF0083: mov var_84, edx
  loc_00DF0089: mov [ebx+00000004h], ecx
  loc_00DF008C: mov ecx, esp
  loc_00DF008E: mov [ebx+00000008h], eax
  loc_00DF0091: mov eax, var_90
  loc_00DF0097: mov [ebx+0000000Ch], eax
  loc_00DF009A: mov eax, var_8C
  loc_00DF00A0: mov [ecx], eax
  loc_00DF00A2: mov eax, var_88
  loc_00DF00A8: mov [ecx+00000004h], eax
  loc_00DF00AB: lea eax, var_4C
  loc_00DF00AE: push eax
  loc_00DF00AF: mov [ecx+00000008h], edx
  loc_00DF00B2: mov edx, var_80
  loc_00DF00B5: mov [ecx+0000000Ch], edx
  loc_00DF00B8: call [004011D0h] ; __vbaI4Var
  loc_00DF00BE: mov var_F0, eax
  loc_00DF00C4: fild real4 ptr var_F0
  loc_00DF00CA: fstp real4 ptr var_F4
  loc_00DF00D0: mov ecx, var_F4
  loc_00DF00D6: push ecx
  loc_00DF00D7: push esi
  loc_00DF00D8: call [edi+00000374h]
  loc_00DF00DE: test eax, eax
  loc_00DF00E0: fnclex
  loc_00DF00E2: jge 00DF00F6h
  loc_00DF00E4: push 00000374h
  loc_00DF00E9: push 006DCB00h
  loc_00DF00EE: push esi
  loc_00DF00EF: push eax
  loc_00DF00F0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF00F6: fld real4 ptr var_C0
  loc_00DF00FC: call [00401208h] ; __vbaFpI4
  loc_00DF0102: lea ecx, var_4C
  loc_00DF0105: mov [esi+00000178h], eax
  loc_00DF010B: call [00401024h] ; __vbaFreeVar
  loc_00DF0111: mov eax, [esi+0000005Ch]
  loc_00DF0114: test eax, eax
  loc_00DF0116: jnz 00DF0137h
  loc_00DF0118: cmp [esi+00000060h], 0000h
  loc_00DF011D: jnz 00DF0137h
  loc_00DF011F: mov edx, [esi+000001D0h]
  loc_00DF0125: mov ecx, [esi+000001D4h]
  loc_00DF012B: lea eax, [esi+0000013Ch]
  loc_00DF0131: push edx
  loc_00DF0132: sub ecx, 00000008h
  loc_00DF0135: jmp 00DF0175h
  loc_00DF0137: mov eax, [esi+00000164h]
  loc_00DF013D: cmp eax, 00000002h
  loc_00DF0140: jz 00DF015Fh
  loc_00DF0142: cmp eax, 00000003h
  loc_00DF0145: jz 00DF015Fh
  loc_00DF0147: mov edx, [esi+000001D0h]
  loc_00DF014D: mov ecx, [esi+000001D4h]
  loc_00DF0153: lea eax, [esi+0000013Ch]
  loc_00DF0159: push edx
  loc_00DF015A: sub ecx, 00000010h
  loc_00DF015D: jmp 00DF0175h
  loc_00DF015F: mov edx, [esi+000001D0h]
  loc_00DF0165: mov ecx, [esi+000001D4h]
  loc_00DF016B: lea eax, [esi+0000013Ch]
  loc_00DF0171: push edx
  loc_00DF0172: sub ecx, 00000018h
  loc_00DF0175: jo 00DF0DE1h
  loc_00DF017B: push ecx
  loc_00DF017C: push 00000000h
  loc_00DF017E: push 00000000h
  loc_00DF0180: push eax
  loc_00DF0181: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF0186: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00DF018C: call ebx
  loc_00DF018E: mov eax, [esi+00000078h]
  loc_00DF0191: test eax, eax
  loc_00DF0193: jz 00DF027Dh
  loc_00DF0199: mov edx, [esi]
  loc_00DF019B: lea eax, var_C0
  loc_00DF01A1: push eax
  loc_00DF01A2: push esi
  loc_00DF01A3: call [edx+000000D8h]
  loc_00DF01A9: test eax, eax
  loc_00DF01AB: fnclex
  loc_00DF01AD: jge 00DF01C1h
  loc_00DF01AF: push 000000D8h
  loc_00DF01B4: push 006DCB00h
  loc_00DF01B9: push esi
  loc_00DF01BA: push eax
  loc_00DF01BB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF01C1: mov ecx, [esi+00000080h]
  loc_00DF01C7: push ecx
  loc_00DF01C8: call [00401190h] ; VarPtr
  loc_00DF01CE: mov var_C4, eax
  loc_00DF01D4: lea edx, [esi+00000138h]
  loc_00DF01DA: lea eax, var_5C
  loc_00DF01DD: mov var_84, edx
  loc_00DF01E3: lea ecx, var_4C
  loc_00DF01E6: push eax
  loc_00DF01E7: lea edx, var_8C
  loc_00DF01ED: push ecx
  loc_00DF01EE: lea eax, var_6C
  loc_00DF01F1: mov edi, 00000003h
  loc_00DF01F6: push edx
  loc_00DF01F7: push eax
  loc_00DF01F8: mov var_B4, 00000410h
  loc_00DF0202: mov var_BC, edi
  loc_00DF0208: mov var_54, 00000000h
  loc_00DF020F: mov var_5C, 00000002h
  loc_00DF0216: mov var_44, 00020000h
  loc_00DF021D: mov var_4C, edi
  loc_00DF0220: mov var_8C, 0000400Bh
  loc_00DF022A: call [004011B4h] ; rtcImmediateIf
  loc_00DF0230: lea ecx, var_BC
  loc_00DF0236: lea edx, var_6C
  loc_00DF0239: push ecx
  loc_00DF023A: lea eax, var_7C
  loc_00DF023D: push edx
  loc_00DF023E: push eax
  loc_00DF023F: call [00401118h] ; __vbaVarOr
  loc_00DF0245: push eax
  loc_00DF0246: call [004011D0h] ; __vbaI4Var
  loc_00DF024C: mov ecx, var_C4
  loc_00DF0252: mov edx, var_C0
  loc_00DF0258: push eax
  loc_00DF0259: lea eax, [esi+0000013Ch]
  loc_00DF025F: push eax
  loc_00DF0260: push FFFFFFFFh
  loc_00DF0262: push ecx
  loc_00DF0263: push edx
  loc_00DF0264: call 006DE300h ; DrawTextW()
  loc_00DF0269: call ebx
  loc_00DF026B: lea eax, var_6C
  loc_00DF026E: lea ecx, var_5C
  loc_00DF0271: push eax
  loc_00DF0272: lea edx, var_4C
  loc_00DF0275: push ecx
  loc_00DF0276: push edx
  loc_00DF0277: push edi
  loc_00DF0278: jmp 00DF036Bh
  loc_00DF027D: mov eax, [esi]
  loc_00DF027F: lea ecx, var_C0
  loc_00DF0285: push ecx
  loc_00DF0286: push esi
  loc_00DF0287: call [eax+000000D8h]
  loc_00DF028D: test eax, eax
  loc_00DF028F: fnclex
  loc_00DF0291: jge 00DF02A5h
  loc_00DF0293: push 000000D8h
  loc_00DF0298: push 006DCB00h
  loc_00DF029D: push esi
  loc_00DF029E: push eax
  loc_00DF029F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF02A5: mov eax, 00000003h
  loc_00DF02AA: lea edx, [esi+00000138h]
  loc_00DF02B0: mov var_BC, eax
  loc_00DF02B6: mov var_4C, eax
  loc_00DF02B9: lea eax, var_5C
  loc_00DF02BC: mov var_84, edx
  loc_00DF02C2: lea ecx, var_4C
  loc_00DF02C5: push eax
  loc_00DF02C6: lea edx, var_8C
  loc_00DF02CC: push ecx
  loc_00DF02CD: lea eax, var_6C
  loc_00DF02D0: push edx
  loc_00DF02D1: push eax
  loc_00DF02D2: mov var_B4, 00000410h
  loc_00DF02DC: mov var_54, 00000000h
  loc_00DF02E3: mov var_5C, 00000002h
  loc_00DF02EA: mov var_44, 00020000h
  loc_00DF02F1: mov var_8C, 0000400Bh
  loc_00DF02FB: call [004011B4h] ; rtcImmediateIf
  loc_00DF0301: lea ecx, var_BC
  loc_00DF0307: lea edx, var_6C
  loc_00DF030A: push ecx
  loc_00DF030B: lea eax, var_7C
  loc_00DF030E: push edx
  loc_00DF030F: push eax
  loc_00DF0310: lea edi, [esi+00000080h]
  loc_00DF0316: call [00401118h] ; __vbaVarOr
  loc_00DF031C: push eax
  loc_00DF031D: call [004011D0h] ; __vbaI4Var
  loc_00DF0323: mov ecx, [edi]
  loc_00DF0325: push eax
  loc_00DF0326: lea eax, [esi+0000013Ch]
  loc_00DF032C: lea edx, var_3C
  loc_00DF032F: push eax
  loc_00DF0330: push FFFFFFFFh
  loc_00DF0332: push ecx
  loc_00DF0333: push edx
  loc_00DF0334: call [004011ECh] ; __vbaStrToAnsi
  loc_00DF033A: push eax
  loc_00DF033B: mov eax, var_C0
  loc_00DF0341: push eax
  loc_00DF0342: call 006DE2BCh ; DrawText(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF0347: call ebx
  loc_00DF0349: mov ecx, var_3C
  loc_00DF034C: push ecx
  loc_00DF034D: push edi
  loc_00DF034E: call [00401160h] ; __vbaStrToUnicode
  loc_00DF0354: lea ecx, var_3C
  loc_00DF0357: call [00401258h] ; __vbaFreeStr
  loc_00DF035D: lea edx, var_6C
  loc_00DF0360: lea eax, var_5C
  loc_00DF0363: push edx
  loc_00DF0364: lea ecx, var_4C
  loc_00DF0367: push eax
  loc_00DF0368: push ecx
  loc_00DF0369: push 00000003h
  loc_00DF036B: call [00401038h] ; __vbaFreeVarList
  loc_00DF0371: add esp, 00000010h
  loc_00DF0374: lea eax, [esi+0000013Ch]
  loc_00DF037A: lea edx, var_34
  loc_00DF037D: push eax
  loc_00DF037E: push edx
  loc_00DF037F: call 006DD98Ch ; CopyRect(%x1v, %x2v)
  loc_00DF0384: call ebx
  loc_00DF0386: mov eax, [esi+00000084h]
  loc_00DF038C: sub eax, 00000000h
  loc_00DF038F: jz 00DF046Eh
  loc_00DF0395: dec eax
  loc_00DF0396: jz 00DF03DAh
  loc_00DF0398: dec eax
  loc_00DF0399: jnz 00DF0492h
  loc_00DF039F: mov eax, [esi+000001D0h]
  loc_00DF03A5: mov ecx, var_28
  loc_00DF03A8: sub eax, ecx
  loc_00DF03AA: mov edi, var_2C
  loc_00DF03AD: jo 00DF0DE1h
  loc_00DF03B3: cdq
  loc_00DF03B4: sub eax, edx
  loc_00DF03B6: lea ecx, var_34
  loc_00DF03B9: sar eax, 01h
  loc_00DF03BB: push eax
  loc_00DF03BC: mov eax, [esi+000001D4h]
  loc_00DF03C2: sub eax, edi
  loc_00DF03C4: jo 00DF0DE1h
  loc_00DF03CA: sub eax, 00000004h
  loc_00DF03CD: jo 00DF0DE1h
  loc_00DF03D3: push eax
  loc_00DF03D4: push ecx
  loc_00DF03D5: jmp 00DF048Bh
  loc_00DF03DA: mov eax, [esi+000001D0h]
  loc_00DF03E0: mov ecx, var_28
  loc_00DF03E3: sub eax, ecx
  loc_00DF03E5: mov edi, var_2C
  loc_00DF03E8: jo 00DF0DE1h
  loc_00DF03EE: cdq
  loc_00DF03EF: sub eax, edx
  loc_00DF03F1: mov edx, [esi+00000178h]
  loc_00DF03F7: sar eax, 01h
  loc_00DF03F9: push eax
  loc_00DF03FA: mov eax, [esi+000001D4h]
  loc_00DF0400: sub eax, edi
  loc_00DF0402: jo 00DF0DE1h
  loc_00DF0408: add eax, edx
  loc_00DF040A: jo 00DF0DE1h
  loc_00DF0410: cdq
  loc_00DF0411: sub eax, edx
  loc_00DF0413: lea edx, var_34
  loc_00DF0416: sar eax, 01h
  loc_00DF0418: push eax
  loc_00DF0419: push edx
  loc_00DF041A: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF041F: call ebx
  loc_00DF0421: cmp [esi+00000060h], 0000h
  loc_00DF0426: jnz 00DF042Fh
  loc_00DF0428: mov eax, [esi+0000005Ch]
  loc_00DF042B: test eax, eax
  loc_00DF042D: jz 00DF043Eh
  loc_00DF042F: push 00000000h
  loc_00DF0431: lea eax, var_34
  loc_00DF0434: push FFFFFFF8h
  loc_00DF0436: push eax
  loc_00DF0437: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF043C: call ebx
  loc_00DF043E: mov eax, [esi+00000164h]
  loc_00DF0444: cmp eax, 00000007h
  loc_00DF0447: jz 00DF0458h
  loc_00DF0449: cmp eax, 00000008h
  loc_00DF044C: jz 00DF0458h
  loc_00DF044E: cmp eax, 00000006h
  loc_00DF0451: jz 00DF0458h
  loc_00DF0453: cmp eax, 00000005h
  loc_00DF0456: jnz 00DF0492h
  loc_00DF0458: mov eax, [esi+00000178h]
  loc_00DF045E: push 00000000h
  loc_00DF0460: cdq
  loc_00DF0461: sub eax, edx
  loc_00DF0463: lea ecx, var_34
  loc_00DF0466: sar eax, 01h
  loc_00DF0468: neg eax
  loc_00DF046A: push eax
  loc_00DF046B: push ecx
  loc_00DF046C: jmp 00DF048Bh
  loc_00DF046E: mov eax, [esi+000001D0h]
  loc_00DF0474: mov edx, var_28
  loc_00DF0477: sub eax, edx
  loc_00DF0479: jo 00DF0DE1h
  loc_00DF047F: cdq
  loc_00DF0480: sub eax, edx
  loc_00DF0482: lea edx, var_34
  loc_00DF0485: sar eax, 01h
  loc_00DF0487: push eax
  loc_00DF0488: push 00000002h
  loc_00DF048A: push edx
  loc_00DF048B: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0490: call ebx
  loc_00DF0492: mov eax, [esi+0000014Ch]
  loc_00DF0498: push 00000000h
  loc_00DF049A: push eax
  loc_00DF049B: call [0040114Ch] ; __vbaObjIs
  loc_00DF04A1: test ax, ax
  loc_00DF04A4: jnz 00DF05BAh
  loc_00DF04AA: mov eax, [esi+00000164h]
  loc_00DF04B0: cmp eax, 00000008h
  loc_00DF04B3: ja 00DF05BAh
  loc_00DF04B9: jmp [eax*4+00DF0DB8h]
  loc_00DF04C0: cmp [esi+00000084h], 00000001h
  loc_00DF04C7: jz 00DF05BAh
  loc_00DF04CD: mov ecx, [esi+00000178h]
  loc_00DF04D3: mov eax, var_34
  loc_00DF04D6: mov edx, ecx
  loc_00DF04D8: add edx, 00000004h
  loc_00DF04DB: jo 00DF0DE1h
  loc_00DF04E1: cmp eax, edx
  loc_00DF04E3: jge 00DF05BDh
  loc_00DF04E9: mov edi, var_2C
  loc_00DF04EC: mov eax, ecx
  loc_00DF04EE: add eax, 00000004h
  loc_00DF04F1: jo 00DF0DE1h
  loc_00DF04F7: add ecx, edi
  loc_00DF04F9: mov var_34, eax
  loc_00DF04FC: jo 00DF0DE1h
  loc_00DF0502: add ecx, 00000004h
  loc_00DF0505: jo 00DF0DE1h
  loc_00DF050B: mov var_2C, ecx
  loc_00DF050E: jmp 00DF05BDh
  loc_00DF0513: mov eax, [esi+000001D4h]
  loc_00DF0519: mov ecx, [esi+00000178h]
  loc_00DF051F: mov edi, var_2C
  loc_00DF0522: mov edx, eax
  loc_00DF0524: sub edx, ecx
  loc_00DF0526: jo 00DF0DE1h
  loc_00DF052C: sub edx, 00000004h
  loc_00DF052F: jo 00DF0DE1h
  loc_00DF0535: cmp edi, edx
  loc_00DF0537: jle 00DF0566h
  loc_00DF0539: sub eax, ecx
  loc_00DF053B: jo 00DF0DE1h
  loc_00DF0541: sub eax, 00000004h
  loc_00DF0544: jo 00DF0DE1h
  loc_00DF054A: mov var_2C, eax
  loc_00DF054D: mov eax, var_34
  loc_00DF0550: sub eax, ecx
  loc_00DF0552: jo 00DF0DE1h
  loc_00DF0558: sub eax, 00000004h
  loc_00DF055B: jo 00DF0DE1h
  loc_00DF0561: mov var_34, eax
  loc_00DF0564: jmp 00DF0569h
  loc_00DF0566: mov eax, var_34
  loc_00DF0569: cmp [esi+00000084h], 00000001h
  loc_00DF0570: jnz 00DF05BDh
  loc_00DF0572: push 00000000h
  loc_00DF0574: push FFFFFFF4h
  loc_00DF0576: jmp 00DF05AFh
  loc_00DF0578: mov eax, [esi+00000174h]
  loc_00DF057E: lea ecx, var_34
  loc_00DF0581: cdq
  loc_00DF0582: sub eax, edx
  loc_00DF0584: sar eax, 01h
  loc_00DF0586: push eax
  loc_00DF0587: push 00000000h
  loc_00DF0589: push ecx
  loc_00DF058A: jmp 00DF05B3h
  loc_00DF058C: mov eax, [esi+00000174h]
  loc_00DF0592: cdq
  loc_00DF0593: sub eax, edx
  loc_00DF0595: lea edx, var_34
  loc_00DF0598: sar eax, 01h
  loc_00DF059A: neg eax
  loc_00DF059C: push eax
  loc_00DF059D: push 00000000h
  loc_00DF059F: push edx
  loc_00DF05A0: jmp 00DF05B3h
  loc_00DF05A2: cmp [esi+00000084h], 00000001h
  loc_00DF05A9: jnz 00DF05BAh
  loc_00DF05AB: push 00000000h
  loc_00DF05AD: push FFFFFFF0h
  loc_00DF05AF: lea eax, var_34
  loc_00DF05B2: push eax
  loc_00DF05B3: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF05B8: call ebx
  loc_00DF05BA: mov eax, var_34
  loc_00DF05BD: cmp [esi+00000084h], 00000002h
  loc_00DF05C4: jnz 00DF05E6h
  loc_00DF05C6: cmp [esi+00000060h], 0000h
  loc_00DF05CB: jnz 00DF05D4h
  loc_00DF05CD: mov ecx, [esi+0000005Ch]
  loc_00DF05D0: test ecx, ecx
  loc_00DF05D2: jz 00DF05E6h
  loc_00DF05D4: push 00000000h
  loc_00DF05D6: lea ecx, var_34
  loc_00DF05D9: push FFFFFFF0h
  loc_00DF05DB: push ecx
  loc_00DF05DC: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF05E1: call ebx
  loc_00DF05E3: mov eax, var_34
  loc_00DF05E6: cmp [esi+00000044h], 0000000Dh
  loc_00DF05EA: jnz 00DF0775h
  loc_00DF05F0: cmp eax, 00000004h
  loc_00DF05F3: jge 00DF05FCh
  loc_00DF05F5: mov var_34, 00000004h
  loc_00DF05FC: mov edx, [esi]
  loc_00DF05FE: lea eax, var_C0
  loc_00DF0604: push eax
  loc_00DF0605: push esi
  loc_00DF0606: call [edx+00000100h]
  loc_00DF060C: test eax, eax
  loc_00DF060E: fnclex
  loc_00DF0610: jge 00DF0624h
  loc_00DF0612: push 00000100h
  loc_00DF0617: push 006DCB00h
  loc_00DF061C: push esi
  loc_00DF061D: push eax
  loc_00DF061E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF0624: fild real4 ptr var_2C
  loc_00DF0627: fstp real8 ptr var_FC
  loc_00DF062D: fld real4 ptr var_C0
  loc_00DF0633: fsub st0, real4 ptr [004013D0h]
  loc_00DF0639: fnstsw ax
  loc_00DF063B: test al, 0Dh
  loc_00DF063D: jnz 00DF0DDCh
  loc_00DF0643: call [004010D8h] ; __vbaFpR8
  loc_00DF0649: fcomp real8 ptr var_FC
  loc_00DF064F: fnstsw ax
  loc_00DF0651: test ah, 01h
  loc_00DF0654: jz 00DF065Dh
  loc_00DF0656: mov eax, 00000001h
  loc_00DF065B: jmp 00DF065Fh
  loc_00DF065D: xor eax, eax
  loc_00DF065F: neg eax
  loc_00DF0661: test ax, ax
  loc_00DF0664: jz 00DF06B1h
  loc_00DF0666: mov ecx, [esi]
  loc_00DF0668: lea edx, var_C0
  loc_00DF066E: push edx
  loc_00DF066F: push esi
  loc_00DF0670: call [ecx+00000100h]
  loc_00DF0676: test eax, eax
  loc_00DF0678: fnclex
  loc_00DF067A: jge 00DF068Eh
  loc_00DF067C: push 00000100h
  loc_00DF0681: push 006DCB00h
  loc_00DF0686: push esi
  loc_00DF0687: push eax
  loc_00DF0688: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF068E: fld real4 ptr var_C0
  loc_00DF0694: fsub st0, real4 ptr [004013D0h]
  loc_00DF069A: mov edi, [00401208h] ; __vbaFpI4
  loc_00DF06A0: fnstsw ax
  loc_00DF06A2: test al, 0Dh
  loc_00DF06A4: jnz 00DF0DDCh
  loc_00DF06AA: call edi
  loc_00DF06AC: mov var_2C, eax
  loc_00DF06AF: jmp 00DF06B7h
  loc_00DF06B1: mov edi, [00401208h] ; __vbaFpI4
  loc_00DF06B7: mov ecx, var_30
  loc_00DF06BA: mov eax, 00000004h
  loc_00DF06BF: cmp ecx, eax
  loc_00DF06C1: jge 00DF06C6h
  loc_00DF06C3: mov var_30, eax
  loc_00DF06C6: mov eax, [esi]
  loc_00DF06C8: lea ecx, var_C0
  loc_00DF06CE: push ecx
  loc_00DF06CF: push esi
  loc_00DF06D0: call [eax+00000108h]
  loc_00DF06D6: test eax, eax
  loc_00DF06D8: fnclex
  loc_00DF06DA: jge 00DF06EEh
  loc_00DF06DC: push 00000108h
  loc_00DF06E1: push 006DCB00h
  loc_00DF06E6: push esi
  loc_00DF06E7: push eax
  loc_00DF06E8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF06EE: fild real4 ptr var_28
  loc_00DF06F1: fstp real8 ptr var_104
  loc_00DF06F7: fld real4 ptr var_C0
  loc_00DF06FD: fsub st0, real4 ptr [004013D0h]
  loc_00DF0703: fnstsw ax
  loc_00DF0705: test al, 0Dh
  loc_00DF0707: jnz 00DF0DDCh
  loc_00DF070D: call [004010D8h] ; __vbaFpR8
  loc_00DF0713: fcomp real8 ptr var_104
  loc_00DF0719: fnstsw ax
  loc_00DF071B: test ah, 01h
  loc_00DF071E: jz 00DF0727h
  loc_00DF0720: mov eax, 00000001h
  loc_00DF0725: jmp 00DF0729h
  loc_00DF0727: xor eax, eax
  loc_00DF0729: neg eax
  loc_00DF072B: test ax, ax
  loc_00DF072E: jz 00DF077Bh
  loc_00DF0730: mov edx, [esi]
  loc_00DF0732: lea eax, var_C0
  loc_00DF0738: push eax
  loc_00DF0739: push esi
  loc_00DF073A: call [edx+00000108h]
  loc_00DF0740: test eax, eax
  loc_00DF0742: fnclex
  loc_00DF0744: jge 00DF0758h
  loc_00DF0746: push 00000108h
  loc_00DF074B: push 006DCB00h
  loc_00DF0750: push esi
  loc_00DF0751: push eax
  loc_00DF0752: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF0758: fld real4 ptr var_C0
  loc_00DF075E: fsub st0, real4 ptr [004013D0h]
  loc_00DF0764: fnstsw ax
  loc_00DF0766: test al, 0Dh
  loc_00DF0768: jnz 00DF0DDCh
  loc_00DF076E: call edi
  loc_00DF0770: mov var_28, eax
  loc_00DF0773: jmp 00DF077Bh
  loc_00DF0775: mov edi, [00401208h] ; __vbaFpI4
  loc_00DF077B: lea ecx, var_34
  loc_00DF077E: lea eax, [esi+0000013Ch]
  loc_00DF0784: push ecx
  loc_00DF0785: push eax
  loc_00DF0786: call 006DD98Ch ; CopyRect(%x1v, %x2v)
  loc_00DF078B: call ebx
  loc_00DF078D: mov edx, [esi]
  loc_00DF078F: push esi
  loc_00DF0790: call [edx+00000924h]
  loc_00DF0796: mov eax, [esi+0000005Ch]
  loc_00DF0799: test eax, eax
  loc_00DF079B: jz 00DF0854h
  loc_00DF07A1: fild real4 ptr [esi+000001D0h]
  loc_00DF07A7: fstp real8 ptr var_10C
  loc_00DF07AD: fld real8 ptr var_10C
  loc_00DF07B3: cmp [00E3D000h], 00000000h
  loc_00DF07BA: jnz 00DF07C4h
  loc_00DF07BC: fdiv st0, real8 ptr [00401338h]
  loc_00DF07C2: jmp 00DF07D5h
  loc_00DF07C4: push [0040133Ch]
  loc_00DF07CA: push [00401338h]
  loc_00DF07D0: call 00402854h ; _adj_fdiv_m64
  loc_00DF07D5: fadd st0, real8 ptr [004013C8h]
  loc_00DF07DB: fnstsw ax
  loc_00DF07DD: test al, 0Dh
  loc_00DF07DF: jnz 00DF0DDCh
  loc_00DF07E5: call edi
  loc_00DF07E7: fild real4 ptr [esi+000001D0h]
  loc_00DF07ED: push eax
  loc_00DF07EE: mov eax, [esi+000001D4h]
  loc_00DF07F4: push eax
  loc_00DF07F5: fstp real8 ptr var_114
  loc_00DF07FB: fld real8 ptr var_114
  loc_00DF0801: cmp [00E3D000h], 00000000h
  loc_00DF0808: jnz 00DF0812h
  loc_00DF080A: fdiv st0, real8 ptr [00401338h]
  loc_00DF0810: jmp 00DF0823h
  loc_00DF0812: push [0040133Ch]
  loc_00DF0818: push [00401338h]
  loc_00DF081E: call 00402854h ; _adj_fdiv_m64
  loc_00DF0823: fsub st0, real8 ptr [004013C0h]
  loc_00DF0829: fnstsw ax
  loc_00DF082B: test al, 0Dh
  loc_00DF082D: jnz 00DF0DDCh
  loc_00DF0833: call edi
  loc_00DF0835: mov ecx, [esi+000001D4h]
  loc_00DF083B: push eax
  loc_00DF083C: sub ecx, 0000000Fh
  loc_00DF083F: lea edx, [esi+00000128h]
  loc_00DF0845: jo 00DF0DE1h
  loc_00DF084B: push ecx
  loc_00DF084C: push edx
  loc_00DF084D: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF0852: call ebx
  loc_00DF0854: xor edi, edi
  loc_00DF0856: cmp [esi+00000060h], di
  loc_00DF085A: jz 00DF0A41h
  loc_00DF0860: mov eax, [esi]
  loc_00DF0862: lea ecx, var_C0
  loc_00DF0868: push ecx
  loc_00DF0869: push esi
  loc_00DF086A: call [eax+00000100h]
  loc_00DF0870: cmp eax, edi
  loc_00DF0872: fnclex
  loc_00DF0874: jge 00DF0888h
  loc_00DF0876: push 00000100h
  loc_00DF087B: push 006DCB00h
  loc_00DF0880: push esi
  loc_00DF0881: push eax
  loc_00DF0882: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF0888: fild real4 ptr [esi+00000144h]
  loc_00DF088E: fstp real8 ptr var_11C
  loc_00DF0894: fld real4 ptr var_C0
  loc_00DF089A: fsub st0, real4 ptr [004013B8h]
  loc_00DF08A0: fnstsw ax
  loc_00DF08A2: test al, 0Dh
  loc_00DF08A4: jnz 00DF0DDCh
  loc_00DF08AA: call [004010D8h] ; __vbaFpR8
  loc_00DF08B0: fcomp real8 ptr var_11C
  loc_00DF08B6: fnstsw ax
  loc_00DF08B8: test ah, 41h
  loc_00DF08BB: jnz 00DF08C4h
  loc_00DF08BD: mov eax, 00000001h
  loc_00DF08C2: jmp 00DF08C6h
  loc_00DF08C4: xor eax, eax
  loc_00DF08C6: neg eax
  loc_00DF08C8: test ax, ax
  loc_00DF08CB: jz 00DF0A41h
  loc_00DF08D1: mov edx, [esi]
  loc_00DF08D3: lea eax, var_C0
  loc_00DF08D9: push eax
  loc_00DF08DA: push esi
  loc_00DF08DB: call [edx+000000D8h]
  loc_00DF08E1: cmp eax, edi
  loc_00DF08E3: fnclex
  loc_00DF08E5: jge 00DF08F9h
  loc_00DF08E7: push 000000D8h
  loc_00DF08EC: push 006DCB00h
  loc_00DF08F1: push esi
  loc_00DF08F2: push eax
  loc_00DF08F3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF08F9: mov ecx, var_C0
  loc_00DF08FF: push 00000007h
  loc_00DF0901: push 00000007h
  loc_00DF0903: push ecx
  loc_00DF0904: call 006DD68Ch ; GetPixel(%x1v, %x2v, %x3v)
  loc_00DF0909: mov var_C4, eax
  loc_00DF090F: call ebx
  loc_00DF0911: fld real8 ptr [004013B0h]
  loc_00DF0917: mov edx, var_C4
  loc_00DF091D: lea ecx, var_D0
  loc_00DF0923: fstp real4 ptr var_CC
  loc_00DF0929: fnstsw ax
  loc_00DF092B: test al, 0Dh
  loc_00DF092D: jnz 00DF0DDCh
  loc_00DF0933: mov var_C8, edx
  loc_00DF0939: push ecx
  loc_00DF093A: lea edx, var_CC
  loc_00DF0940: lea ecx, var_C8
  loc_00DF0946: mov eax, [esi]
  loc_00DF0948: push edx
  loc_00DF0949: push ecx
  loc_00DF094A: push esi
  loc_00DF094B: call [eax+0000097Ch]
  loc_00DF0951: mov ecx, var_D0
  loc_00DF0957: mov eax, [esi+000001D4h]
  loc_00DF095D: push ecx
  loc_00DF095E: mov ecx, [esi+000001D0h]
  loc_00DF0964: sub ecx, 00000003h
  loc_00DF0967: mov edx, [esi]
  loc_00DF0969: jo 00DF0DE1h
  loc_00DF096F: push ecx
  loc_00DF0970: mov ecx, eax
  loc_00DF0972: sub ecx, 00000010h
  loc_00DF0975: jo 00DF0DE1h
  loc_00DF097B: sub eax, 00000010h
  loc_00DF097E: push ecx
  loc_00DF097F: push 00000003h
  loc_00DF0981: jo 00DF0DE1h
  loc_00DF0987: push eax
  loc_00DF0988: push esi
  loc_00DF0989: call [edx+000008E4h]
  loc_00DF098F: mov edx, [esi]
  loc_00DF0991: lea eax, var_C0
  loc_00DF0997: push eax
  loc_00DF0998: push esi
  loc_00DF0999: call [edx+000000D8h]
  loc_00DF099F: cmp eax, edi
  loc_00DF09A1: fnclex
  loc_00DF09A3: jge 00DF09B7h
  loc_00DF09A5: push 000000D8h
  loc_00DF09AA: push 006DCB00h
  loc_00DF09AF: push esi
  loc_00DF09B0: push eax
  loc_00DF09B1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF09B7: mov ecx, var_C0
  loc_00DF09BD: push 00000007h
  loc_00DF09BF: push 00000007h
  loc_00DF09C1: push ecx
  loc_00DF09C2: call 006DD68Ch ; GetPixel(%x1v, %x2v, %x3v)
  loc_00DF09C7: mov var_C4, eax
  loc_00DF09CD: call ebx
  loc_00DF09CF: mov edx, var_C4
  loc_00DF09D5: mov eax, [esi]
  loc_00DF09D7: lea ecx, var_D0
  loc_00DF09DD: mov var_C8, edx
  loc_00DF09E3: push ecx
  loc_00DF09E4: lea edx, var_CC
  loc_00DF09EA: lea ecx, var_C8
  loc_00DF09F0: push edx
  loc_00DF09F1: push ecx
  loc_00DF09F2: push esi
  loc_00DF09F3: mov var_CC, 3DCCCCCDh
  loc_00DF09FD: call [eax+0000097Ch]
  loc_00DF0A03: mov ecx, var_D0
  loc_00DF0A09: mov eax, [esi+000001D4h]
  loc_00DF0A0F: push ecx
  loc_00DF0A10: mov ecx, [esi+000001D0h]
  loc_00DF0A16: sub ecx, 00000003h
  loc_00DF0A19: mov edx, [esi]
  loc_00DF0A1B: jo 00DF0DE1h
  loc_00DF0A21: push ecx
  loc_00DF0A22: mov ecx, eax
  loc_00DF0A24: sub ecx, 0000000Fh
  loc_00DF0A27: jo 00DF0DE1h
  loc_00DF0A2D: sub eax, 0000000Fh
  loc_00DF0A30: push ecx
  loc_00DF0A31: push 00000003h
  loc_00DF0A33: jo 00DF0DE1h
  loc_00DF0A39: push eax
  loc_00DF0A3A: push esi
  loc_00DF0A3B: call [edx+000008E4h]
  loc_00DF0A41: cmp [esi+00000048h], 00000002h
  loc_00DF0A45: jnz 00DF0AA7h
  loc_00DF0A47: mov eax, [esi+00000044h]
  loc_00DF0A4A: cmp eax, 0000000Bh
  loc_00DF0A4D: jz 00DF0A71h
  loc_00DF0A4F: cmp eax, 00000001h
  loc_00DF0A52: jz 00DF0A71h
  loc_00DF0A54: cmp eax, 0000000Ch
  loc_00DF0A57: jz 00DF0A71h
  loc_00DF0A59: cmp eax, 0000000Ah
  loc_00DF0A5C: jz 00DF0A71h
  loc_00DF0A5E: cmp eax, 00000005h
  loc_00DF0A61: jz 00DF0A71h
  loc_00DF0A63: cmp eax, 00000006h
  loc_00DF0A66: jz 00DF0A71h
  loc_00DF0A68: cmp eax, 00000007h
  loc_00DF0A6B: jz 00DF0A71h
  loc_00DF0A6D: cmp eax, edi
  loc_00DF0A6F: jnz 00DF0AA7h
  loc_00DF0A71: push 00000001h
  loc_00DF0A73: lea eax, [esi+0000013Ch]
  loc_00DF0A79: push 00000001h
  loc_00DF0A7B: push eax
  loc_00DF0A7C: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0A81: call ebx
  loc_00DF0A83: push 00000001h
  loc_00DF0A85: lea edx, [esi+000001C0h]
  loc_00DF0A8B: push 00000001h
  loc_00DF0A8D: push edx
  loc_00DF0A8E: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0A93: call ebx
  loc_00DF0A95: push 00000001h
  loc_00DF0A97: lea eax, [esi+00000128h]
  loc_00DF0A9D: push 00000001h
  loc_00DF0A9F: push eax
  loc_00DF0AA0: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0AA5: call ebx
  loc_00DF0AA7: cmp [esi+00000170h], di
  loc_00DF0AAE: jz 00DF0B3Eh
  loc_00DF0AB4: cmp [esi+00000048h], 00000001h
  loc_00DF0AB8: jnz 00DF0B3Eh
  loc_00DF0ABE: mov ecx, [esi]
  loc_00DF0AC0: lea edx, var_C4
  loc_00DF0AC6: lea eax, var_C0
  loc_00DF0ACC: push edx
  loc_00DF0ACD: push eax
  loc_00DF0ACE: push 00C0C0C0h
  loc_00DF0AD3: push esi
  loc_00DF0AD4: mov var_C0, 00000000h
  loc_00DF0ADE: call [ecx+0000090Ch]
  loc_00DF0AE4: mov ecx, var_C4
  loc_00DF0AEA: mov edx, [esi]
  loc_00DF0AEC: lea eax, var_38
  loc_00DF0AEF: lea edi, [esi+000001C0h]
  loc_00DF0AF5: push eax
  loc_00DF0AF6: push edi
  loc_00DF0AF7: push esi
  loc_00DF0AF8: mov var_38, ecx
  loc_00DF0AFB: call [edx+00000928h]
  loc_00DF0B01: lea ecx, var_24
  loc_00DF0B04: push edi
  loc_00DF0B05: push ecx
  loc_00DF0B06: call 006DD98Ch ; CopyRect(%x1v, %x2v)
  loc_00DF0B0B: call ebx
  loc_00DF0B0D: push FFFFFFFEh
  loc_00DF0B0F: lea edx, var_24
  loc_00DF0B12: push FFFFFFFEh
  loc_00DF0B14: push edx
  loc_00DF0B15: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0B1A: call ebx
  loc_00DF0B1C: mov eax, [esi]
  loc_00DF0B1E: lea ecx, var_C0
  loc_00DF0B24: lea edx, var_24
  loc_00DF0B27: push ecx
  loc_00DF0B28: push edx
  loc_00DF0B29: push esi
  loc_00DF0B2A: mov var_C0, FFFFFFFFh
  loc_00DF0B34: call [eax+00000928h]
  loc_00DF0B3A: xor edi, edi
  loc_00DF0B3C: jmp 00DF0B5Fh
  loc_00DF0B3E: mov eax, [esi]
  loc_00DF0B40: lea ecx, var_C0
  loc_00DF0B46: lea edx, [esi+000001C0h]
  loc_00DF0B4C: push ecx
  loc_00DF0B4D: push edx
  loc_00DF0B4E: push esi
  loc_00DF0B4F: mov var_C0, FFFFFFFFh
  loc_00DF0B59: call [eax+00000928h]
  loc_00DF0B5F: cmp [esi+00000160h], di
  loc_00DF0B66: jz 00DF0B80h
  loc_00DF0B68: cmp [esi+00000170h], di
  loc_00DF0B6F: jz 00DF0B77h
  loc_00DF0B71: cmp [esi+00000048h], 00000001h
  loc_00DF0B75: jz 00DF0B80h
  loc_00DF0B77: mov eax, [esi]
  loc_00DF0B79: push esi
  loc_00DF0B7A: call [eax+0000092Ch]
  loc_00DF0B80: cmp [esi+00000068h], edi
  loc_00DF0B83: jz 00DF0B8Eh
  loc_00DF0B85: mov ecx, [esi]
  loc_00DF0B87: push esi
  loc_00DF0B88: call [ecx+00000930h]
  loc_00DF0B8E: cmp [esi+0000007Ch], di
  loc_00DF0B92: jz 00DF0C54h
  loc_00DF0B98: mov eax, [esi+00000048h]
  loc_00DF0B9B: mov var_C0, edi
  loc_00DF0BA1: cmp eax, 00000001h
  loc_00DF0BA4: jnz 00DF0C06h
  loc_00DF0BA6: mov edx, [esi]
  loc_00DF0BA8: lea eax, var_C4
  loc_00DF0BAE: push eax
  loc_00DF0BAF: mov eax, [esi+00000094h]
  loc_00DF0BB5: lea ecx, var_C0
  loc_00DF0BBB: push ecx
  loc_00DF0BBC: push eax
  loc_00DF0BBD: push esi
  loc_00DF0BBE: call [edx+0000090Ch]
  loc_00DF0BC4: mov ecx, var_C4
  loc_00DF0BCA: mov edx, [esi]
  loc_00DF0BCC: lea eax, var_D0
  loc_00DF0BD2: mov var_C8, ecx
  loc_00DF0BD8: push eax
  loc_00DF0BD9: lea ecx, var_CC
  loc_00DF0BDF: lea eax, var_C8
  loc_00DF0BE5: push ecx
  loc_00DF0BE6: push eax
  loc_00DF0BE7: lea eax, [esi+0000013Ch]
  loc_00DF0BED: push eax
  loc_00DF0BEE: push esi
  loc_00DF0BEF: mov var_D0, edi
  loc_00DF0BF5: mov var_CC, edi
  loc_00DF0BFB: call [edx+00000934h]
  loc_00DF0C01: jmp 00DF0CA0h
  loc_00DF0C06: mov ecx, [esi]
  loc_00DF0C08: lea edx, var_C4
  loc_00DF0C0E: push edx
  loc_00DF0C0F: mov edx, [esi+00000090h]
  loc_00DF0C15: lea eax, var_C0
  loc_00DF0C1B: push eax
  loc_00DF0C1C: push edx
  loc_00DF0C1D: push esi
  loc_00DF0C1E: call [ecx+0000090Ch]
  loc_00DF0C24: mov eax, var_C4
  loc_00DF0C2A: mov ecx, [esi]
  loc_00DF0C2C: mov var_C8, eax
  loc_00DF0C32: lea edx, var_D0
  loc_00DF0C38: lea eax, var_CC
  loc_00DF0C3E: push edx
  loc_00DF0C3F: mov var_D0, edi
  loc_00DF0C45: mov var_CC, edi
  loc_00DF0C4B: push eax
  loc_00DF0C4C: lea edx, var_C8
  loc_00DF0C52: jmp 00DF0C91h
  loc_00DF0C54: push 00000011h
  loc_00DF0C56: call 006DDDB4h ; GetSysColor(%x1v)
  loc_00DF0C5B: mov var_C0, eax
  loc_00DF0C61: call ebx
  loc_00DF0C63: mov eax, var_C0
  loc_00DF0C69: mov ecx, [esi]
  loc_00DF0C6B: mov var_C4, eax
  loc_00DF0C71: lea edx, var_CC
  loc_00DF0C77: lea eax, var_C8
  loc_00DF0C7D: push edx
  loc_00DF0C7E: mov var_CC, edi
  loc_00DF0C84: mov var_C8, edi
  loc_00DF0C8A: push eax
  loc_00DF0C8B: lea edx, var_C4
  loc_00DF0C91: lea eax, [esi+0000013Ch]
  loc_00DF0C97: push edx
  loc_00DF0C98: push eax
  loc_00DF0C99: push esi
  loc_00DF0C9A: call [ecx+00000934h]
  loc_00DF0CA0: cmp [esi+0000005Ch], edi
  loc_00DF0CA3: jz 00DF0D72h
  loc_00DF0CA9: mov eax, [esi+00000044h]
  loc_00DF0CAC: cmp eax, edi
  loc_00DF0CAE: jz 00DF0CC9h
  loc_00DF0CB0: cmp eax, 0000000Bh
  loc_00DF0CB3: jz 00DF0CC9h
  loc_00DF0CB5: cmp eax, 00000001h
  loc_00DF0CB8: jz 00DF0CC9h
  loc_00DF0CBA: cmp eax, 0000000Ch
  loc_00DF0CBD: jz 00DF0CC9h
  loc_00DF0CBF: cmp eax, 00000007h
  loc_00DF0CC2: jz 00DF0CC9h
  loc_00DF0CC4: cmp eax, 00000006h
  loc_00DF0CC7: jnz 00DF0CE1h
  loc_00DF0CC9: cmp [esi+00000048h], 00000002h
  loc_00DF0CCD: jnz 00DF0CE1h
  loc_00DF0CCF: push 00000001h
  loc_00DF0CD1: lea eax, [esi+00000128h]
  loc_00DF0CD7: push 00000001h
  loc_00DF0CD9: push eax
  loc_00DF0CDA: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0CDF: call ebx
  loc_00DF0CE1: cmp [esi+0000007Ch], di
  loc_00DF0CE5: jz 00DF0D2Eh
  loc_00DF0CE7: mov ecx, [esi]
  loc_00DF0CE9: lea edx, var_C4
  loc_00DF0CEF: push edx
  loc_00DF0CF0: mov edx, [esi+00000090h]
  loc_00DF0CF6: lea eax, var_C0
  loc_00DF0CFC: mov var_C0, edi
  loc_00DF0D02: push eax
  loc_00DF0D03: push edx
  loc_00DF0D04: push esi
  loc_00DF0D05: call [ecx+0000090Ch]
  loc_00DF0D0B: mov eax, [esi+00000010h]
  loc_00DF0D0E: mov edx, var_C4
  loc_00DF0D14: push edx
  loc_00DF0D15: push eax
  loc_00DF0D16: mov ecx, [eax]
  loc_00DF0D18: call [ecx+0000006Ch]
  loc_00DF0D1B: cmp eax, edi
  loc_00DF0D1D: fnclex
  loc_00DF0D1F: jge 00DF0D65h
  loc_00DF0D21: mov ecx, [esi+00000010h]
  loc_00DF0D24: push 0000006Ch
  loc_00DF0D26: push 006DCB00h
  loc_00DF0D2B: push ecx
  loc_00DF0D2C: jmp 00DF0D5Eh
  loc_00DF0D2E: push 00000011h
  loc_00DF0D30: call 006DDDB4h ; GetSysColor(%x1v)
  loc_00DF0D35: mov var_C0, eax
  loc_00DF0D3B: call ebx
  loc_00DF0D3D: mov eax, [esi+00000010h]
  loc_00DF0D40: mov ecx, var_C0
  loc_00DF0D46: push ecx
  loc_00DF0D47: push eax
  loc_00DF0D48: mov edx, [eax]
  loc_00DF0D4A: call [edx+0000006Ch]
  loc_00DF0D4D: cmp eax, edi
  loc_00DF0D4F: fnclex
  loc_00DF0D51: jge 00DF0D65h
  loc_00DF0D53: mov edx, [esi+00000010h]
  loc_00DF0D56: push 0000006Ch
  loc_00DF0D58: push 006DCB00h
  loc_00DF0D5D: push edx
  loc_00DF0D5E: push eax
  loc_00DF0D5F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF0D65: mov ecx, [esi+0000005Ch]
  loc_00DF0D68: mov eax, [esi]
  loc_00DF0D6A: push ecx
  loc_00DF0D6B: push esi
  loc_00DF0D6C: call [eax+00000918h]
  loc_00DF0D72: fwait
  loc_00DF0D73: push 00DF0DA0h
  loc_00DF0D78: jmp 00DF0D9Fh
  loc_00DF0D7A: lea ecx, var_3C
  loc_00DF0D7D: call [00401258h] ; __vbaFreeStr
  loc_00DF0D83: lea edx, var_7C
  loc_00DF0D86: lea eax, var_6C
  loc_00DF0D89: push edx
  loc_00DF0D8A: lea ecx, var_5C
  loc_00DF0D8D: push eax
  loc_00DF0D8E: lea edx, var_4C
  loc_00DF0D91: push ecx
  loc_00DF0D92: push edx
  loc_00DF0D93: push 00000004h
  loc_00DF0D95: call [00401038h] ; __vbaFreeVarList
  loc_00DF0D9B: add esp, 00000014h
  loc_00DF0D9E: ret
  loc_00DF0D9F: ret
  loc_00DF0DA0: mov ecx, var_10
  loc_00DF0DA3: pop edi
  loc_00DF0DA4: pop esi
  loc_00DF0DA5: xor eax, eax
  loc_00DF0DA7: mov fs:[00000000h], ecx
  loc_00DF0DAE: pop ebx
  loc_00DF0DAF: mov esp, ebp
  loc_00DF0DB1: pop ebp
  loc_00DF0DB2: retn 0004h
End Sub

Private Sub Proc_2_116_DF0DF0
  loc_00DF0DF0: push ebp
  loc_00DF0DF1: mov ebp, esp
  loc_00DF0DF3: sub esp, 00000008h
  loc_00DF0DF6: push 00402836h ; __vbaExceptHandler
  loc_00DF0DFB: mov eax, fs:[00000000h]
  loc_00DF0E01: push eax
  loc_00DF0E02: mov fs:[00000000h], esp
  loc_00DF0E09: sub esp, 00000070h
  loc_00DF0E0C: push ebx
  loc_00DF0E0D: push esi
  loc_00DF0E0E: push edi
  loc_00DF0E0F: mov var_8, esp
  loc_00DF0E12: mov var_4, 004013E8h
  loc_00DF0E19: mov esi, Me
  loc_00DF0E1C: xor eax, eax
  loc_00DF0E1E: mov var_20, eax
  loc_00DF0E21: mov var_30, eax
  loc_00DF0E24: mov var_40, eax
  loc_00DF0E27: mov var_50, eax
  loc_00DF0E2A: mov var_60, eax
  loc_00DF0E2D: mov var_70, eax
  loc_00DF0E30: push eax
  loc_00DF0E31: mov eax, [esi+0000014Ch]
  loc_00DF0E37: push eax
  loc_00DF0E38: call [0040114Ch] ; __vbaObjIs
  loc_00DF0E3E: test ax, ax
  loc_00DF0E41: jnz 00DF11CDh
  loc_00DF0E47: lea edx, var_50
  loc_00DF0E4A: lea eax, var_20
  loc_00DF0E4D: lea ecx, [esi+00000080h]
  loc_00DF0E53: push edx
  loc_00DF0E54: push eax
  loc_00DF0E55: lea edi, [esi+000001C0h]
  loc_00DF0E5B: mov var_48, ecx
  loc_00DF0E5E: mov var_50, 00004008h
  loc_00DF0E65: call [004010C8h] ; rtcTrimVar
  loc_00DF0E6B: mov ebx, [esi+00000164h]
  loc_00DF0E71: xor ecx, ecx
  loc_00DF0E73: cmp ebx, 00000004h
  loc_00DF0E76: lea edx, var_20
  loc_00DF0E79: setnz cl
  loc_00DF0E7C: neg ecx
  loc_00DF0E7E: mov var_68, cx
  loc_00DF0E82: lea eax, var_60
  loc_00DF0E85: push edx
  loc_00DF0E86: lea ecx, var_30
  loc_00DF0E89: push eax
  loc_00DF0E8A: push ecx
  loc_00DF0E8B: mov var_58, 006DCC80h
  loc_00DF0E92: mov var_60, 00008008h
  loc_00DF0E99: mov var_70, 0000000Bh
  loc_00DF0EA0: call [00401064h] ; __vbaVarCmpNe
  loc_00DF0EA6: push eax
  loc_00DF0EA7: lea edx, var_70
  loc_00DF0EAA: lea eax, var_40
  loc_00DF0EAD: push edx
  loc_00DF0EAE: push eax
  loc_00DF0EAF: call [00401150h] ; __vbaVarAnd
  loc_00DF0EB5: push eax
  loc_00DF0EB6: call [004010DCh] ; __vbaBoolVarNull
  loc_00DF0EBC: lea ecx, var_70
  loc_00DF0EBF: lea edx, var_20
  loc_00DF0EC2: push ecx
  loc_00DF0EC3: push edx
  loc_00DF0EC4: push 00000002h
  loc_00DF0EC6: mov ebx, eax
  loc_00DF0EC8: call [00401038h] ; __vbaFreeVarList
  loc_00DF0ECE: add esp, 0000000Ch
  loc_00DF0ED1: test bx, bx
  loc_00DF0ED4: jz 00DF115Bh
  loc_00DF0EDA: mov eax, [esi+00000164h]
  loc_00DF0EE0: cmp eax, 00000008h
  loc_00DF0EE3: ja 00DF11AEh
  loc_00DF0EE9: jmp [eax*4+00DF1204h]
  loc_00DF0EF0: mov [edi], 00000003h
  loc_00DF0EF6: mov eax, [esi+000001D0h]
  loc_00DF0EFC: sub eax, [esi+00000174h]
  loc_00DF0F02: jo 00DF1228h
  loc_00DF0F08: cdq
  loc_00DF0F09: sub eax, edx
  loc_00DF0F0B: sar eax, 01h
  loc_00DF0F0D: mov [edi+00000004h], eax
  loc_00DF0F10: mov eax, [edi]
  loc_00DF0F12: test eax, eax
  loc_00DF0F14: jge 00DF11AEh
  loc_00DF0F1A: mov eax, [esi+00000178h]
  loc_00DF0F20: push 00000000h
  loc_00DF0F22: push eax
  loc_00DF0F23: push edi
  loc_00DF0F24: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0F29: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00DF0F2F: call ebx
  loc_00DF0F31: mov ecx, [esi+00000178h]
  loc_00DF0F37: push 00000000h
  loc_00DF0F39: lea edx, [esi+0000013Ch]
  loc_00DF0F3F: push ecx
  loc_00DF0F40: push edx
  loc_00DF0F41: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0F46: call ebx
  loc_00DF0F48: jmp 00DF11AEh
  loc_00DF0F4D: mov eax, [esi+0000013Ch]
  loc_00DF0F53: mov ebx, [esi+00000178h]
  loc_00DF0F59: sub eax, ebx
  loc_00DF0F5B: jo 00DF1228h
  loc_00DF0F61: sub eax, 00000004h
  loc_00DF0F64: jo 00DF1228h
  loc_00DF0F6A: mov [edi], eax
  loc_00DF0F6C: mov eax, [esi+000001D0h]
  loc_00DF0F72: sub eax, [esi+00000174h]
  loc_00DF0F78: jo 00DF1228h
  loc_00DF0F7E: cdq
  loc_00DF0F7F: sub eax, edx
  loc_00DF0F81: sar eax, 01h
  loc_00DF0F83: mov [edi+00000004h], eax
  loc_00DF0F86: jmp 00DF11AEh
  loc_00DF0F8B: mov ecx, [esi+000001D4h]
  loc_00DF0F91: mov ebx, [esi+00000178h]
  loc_00DF0F97: sub ecx, ebx
  loc_00DF0F99: jo 00DF1228h
  loc_00DF0F9F: sub ecx, 00000003h
  loc_00DF0FA2: jo 00DF1228h
  loc_00DF0FA8: mov [edi], ecx
  loc_00DF0FAA: mov eax, [esi+000001D0h]
  loc_00DF0FB0: sub eax, [esi+00000174h]
  loc_00DF0FB6: jo 00DF1228h
  loc_00DF0FBC: cdq
  loc_00DF0FBD: sub eax, edx
  loc_00DF0FBF: sar eax, 01h
  loc_00DF0FC1: mov [edi+00000004h], eax
  loc_00DF0FC4: cmp [esi+00000060h], 0000h
  loc_00DF0FC9: jnz 00DF0FD2h
  loc_00DF0FCB: mov eax, [esi+0000005Ch]
  loc_00DF0FCE: test eax, eax
  loc_00DF0FD0: jz 00DF0FE2h
  loc_00DF0FD2: push 00000000h
  loc_00DF0FD4: push FFFFFFF0h
  loc_00DF0FD6: push edi
  loc_00DF0FD7: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF0FDC: call [00401074h] ; __vbaSetSystemError
  loc_00DF0FE2: mov eax, [esi+00000144h]
  loc_00DF0FE8: mov ecx, [edi]
  loc_00DF0FEA: mov edx, eax
  loc_00DF0FEC: add edx, 00000002h
  loc_00DF0FEF: jo 00DF1228h
  loc_00DF0FF5: cmp ecx, edx
  loc_00DF0FF7: jge 00DF11AEh
  loc_00DF0FFD: add eax, 00000002h
  loc_00DF1000: jo 00DF1228h
  loc_00DF1006: mov [edi], eax
  loc_00DF1008: jmp 00DF11AEh
  loc_00DF100D: mov eax, [esi+00000144h]
  loc_00DF1013: add eax, 00000004h
  loc_00DF1016: jo 00DF1228h
  loc_00DF101C: mov [edi], eax
  loc_00DF101E: mov eax, [esi+000001D0h]
  loc_00DF1024: sub eax, [esi+00000174h]
  loc_00DF102A: jo 00DF1228h
  loc_00DF1030: cdq
  loc_00DF1031: sub eax, edx
  loc_00DF1033: sar eax, 01h
  loc_00DF1035: mov [edi+00000004h], eax
  loc_00DF1038: cmp [esi+00000060h], 0000h
  loc_00DF103D: jnz 00DF1046h
  loc_00DF103F: mov eax, [esi+0000005Ch]
  loc_00DF1042: test eax, eax
  loc_00DF1044: jz 00DF1056h
  loc_00DF1046: push 00000000h
  loc_00DF1048: push FFFFFFF0h
  loc_00DF104A: push edi
  loc_00DF104B: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF1050: call [00401074h] ; __vbaSetSystemError
  loc_00DF1056: mov eax, [esi+00000144h]
  loc_00DF105C: mov edx, [edi]
  loc_00DF105E: mov ecx, eax
  loc_00DF1060: add ecx, 00000002h
  loc_00DF1063: jo 00DF1228h
  loc_00DF1069: cmp edx, ecx
  loc_00DF106B: jge 00DF11AEh
  loc_00DF1071: add eax, 00000002h
  loc_00DF1074: jo 00DF1228h
  loc_00DF107A: mov [edi], eax
  loc_00DF107C: jmp 00DF11AEh
  loc_00DF1081: mov eax, [esi+000001D4h]
  loc_00DF1087: mov ecx, [esi+00000178h]
  loc_00DF108D: sub eax, ecx
  loc_00DF108F: jo 00DF1228h
  loc_00DF1095: cdq
  loc_00DF1096: sub eax, edx
  loc_00DF1098: sar eax, 01h
  loc_00DF109A: mov [edi], eax
  loc_00DF109C: mov edx, [esi+00000140h]
  loc_00DF10A2: sub edx, [esi+00000174h]
  loc_00DF10A8: jo 00DF1228h
  loc_00DF10AE: sub edx, 00000002h
  loc_00DF10B1: jo 00DF1228h
  loc_00DF10B7: mov [edi+00000004h], edx
  loc_00DF10BA: jmp 00DF114Bh
  loc_00DF10BF: mov eax, [esi+000001D4h]
  loc_00DF10C5: mov ecx, [esi+00000178h]
  loc_00DF10CB: sub eax, ecx
  loc_00DF10CD: mov [edi+00000004h], 00000004h
  loc_00DF10D4: jo 00DF1228h
  loc_00DF10DA: cdq
  loc_00DF10DB: sub eax, edx
  loc_00DF10DD: sar eax, 01h
  loc_00DF10DF: mov [edi], eax
  loc_00DF10E1: jmp 00DF114Bh
  loc_00DF10E3: mov eax, [esi+000001D4h]
  loc_00DF10E9: mov ecx, [esi+00000178h]
  loc_00DF10EF: sub eax, ecx
  loc_00DF10F1: jo 00DF1228h
  loc_00DF10F7: cdq
  loc_00DF10F8: sub eax, edx
  loc_00DF10FA: sar eax, 01h
  loc_00DF10FC: mov [edi], eax
  loc_00DF10FE: mov eax, [esi+00000148h]
  loc_00DF1104: add eax, 00000002h
  loc_00DF1107: jo 00DF1228h
  loc_00DF110D: mov [edi+00000004h], eax
  loc_00DF1110: jmp 00DF114Bh
  loc_00DF1112: mov eax, [esi+000001D4h]
  loc_00DF1118: mov ecx, [esi+00000178h]
  loc_00DF111E: sub eax, ecx
  loc_00DF1120: jo 00DF1228h
  loc_00DF1126: cdq
  loc_00DF1127: sub eax, edx
  loc_00DF1129: sar eax, 01h
  loc_00DF112B: mov [edi], eax
  loc_00DF112D: mov ecx, [esi+000001D0h]
  loc_00DF1133: sub ecx, [esi+00000174h]
  loc_00DF1139: jo 00DF1228h
  loc_00DF113F: sub ecx, 00000004h
  loc_00DF1142: jo 00DF1228h
  loc_00DF1148: mov [edi+00000004h], ecx
  loc_00DF114B: cmp [esi+00000060h], 0000h
  loc_00DF1150: jnz 00DF119Eh
  loc_00DF1152: mov eax, [esi+0000005Ch]
  loc_00DF1155: test eax, eax
  loc_00DF1157: jz 00DF11AEh
  loc_00DF1159: jmp 00DF119Eh
  loc_00DF115B: mov eax, [esi+000001D4h]
  loc_00DF1161: mov ecx, [esi+00000178h]
  loc_00DF1167: sub eax, ecx
  loc_00DF1169: jo 00DF1228h
  loc_00DF116F: cdq
  loc_00DF1170: sub eax, edx
  loc_00DF1172: sar eax, 01h
  loc_00DF1174: mov [edi], eax
  loc_00DF1176: mov eax, [esi+000001D0h]
  loc_00DF117C: sub eax, [esi+00000174h]
  loc_00DF1182: jo 00DF1228h
  loc_00DF1188: cdq
  loc_00DF1189: sub eax, edx
  loc_00DF118B: sar eax, 01h
  loc_00DF118D: mov [edi+00000004h], eax
  loc_00DF1190: cmp [esi+00000060h], 0000h
  loc_00DF1195: jnz 00DF119Eh
  loc_00DF1197: mov eax, [esi+0000005Ch]
  loc_00DF119A: test eax, eax
  loc_00DF119C: jz 00DF11AEh
  loc_00DF119E: push 00000000h
  loc_00DF11A0: push FFFFFFF8h
  loc_00DF11A2: push edi
  loc_00DF11A3: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF11A8: call [00401074h] ; __vbaSetSystemError
  loc_00DF11AE: mov edx, [esi+00000178h]
  loc_00DF11B4: mov ebx, [edi]
  loc_00DF11B6: mov ecx, [edi+00000004h]
  loc_00DF11B9: add edx, ebx
  loc_00DF11BB: jo 00DF1228h
  loc_00DF11BD: mov [edi+00000008h], edx
  loc_00DF11C0: mov eax, [esi+00000174h]
  loc_00DF11C6: add eax, ecx
  loc_00DF11C8: jo 00DF1228h
  loc_00DF11CA: mov [edi+0000000Ch], eax
  loc_00DF11CD: push 00DF11EDh
  loc_00DF11D2: jmp 00DF11ECh
  loc_00DF11D4: lea ecx, var_40
  loc_00DF11D7: lea edx, var_30
  loc_00DF11DA: push ecx
  loc_00DF11DB: lea eax, var_20
  loc_00DF11DE: push edx
  loc_00DF11DF: push eax
  loc_00DF11E0: push 00000003h
  loc_00DF11E2: call [00401038h] ; __vbaFreeVarList
  loc_00DF11E8: add esp, 00000010h
  loc_00DF11EB: ret
  loc_00DF11EC: ret
  loc_00DF11ED: mov ecx, var_10
  loc_00DF11F0: pop edi
  loc_00DF11F1: pop esi
  loc_00DF11F2: xor eax, eax
  loc_00DF11F4: mov fs:[00000000h], ecx
  loc_00DF11FB: pop ebx
  loc_00DF11FC: mov esp, ebp
  loc_00DF11FE: pop ebp
  loc_00DF11FF: retn 0004h
End Sub

Private Sub Proc_2_117_DF1230(arg_C, arg_10) 'DF1230
  loc_00DF1230: push ebp
  loc_00DF1231: mov ebp, esp
  loc_00DF1233: sub esp, 00000008h
  loc_00DF1236: push 00402836h ; __vbaExceptHandler
  loc_00DF123B: mov eax, fs:[00000000h]
  loc_00DF1241: push eax
  loc_00DF1242: mov fs:[00000000h], esp
  loc_00DF1249: sub esp, 0000002Ch
  loc_00DF124C: push ebx
  loc_00DF124D: push esi
  loc_00DF124E: push edi
  loc_00DF124F: mov var_8, esp
  loc_00DF1252: mov var_4, 004013F8h
  loc_00DF1259: mov esi, Me
  loc_00DF125C: xor ebx, ebx
  loc_00DF125E: mov var_18, ebx
  loc_00DF1261: mov var_28, ebx
  loc_00DF1264: mov eax, [esi+000001B4h]
  loc_00DF126A: lea edi, [esi+000001B4h]
  loc_00DF1270: cmp eax, ebx
  loc_00DF1272: mov var_2C, ebx
  loc_00DF1275: mov var_30, ebx
  loc_00DF1278: mov var_34, ebx
  loc_00DF127B: jnz 00DF1289h
  loc_00DF127D: push edi
  loc_00DF127E: push 006DEA0Ch
  loc_00DF1283: call [004011A0h] ; __vbaNew2
  loc_00DF1289: mov eax, [edi]
  loc_00DF128B: push ebx
  loc_00DF128C: push 00000003h
  loc_00DF128E: lea ecx, var_28
  loc_00DF1291: push eax
  loc_00DF1292: push ecx
  loc_00DF1293: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DF1299: add esp, 00000010h
  loc_00DF129C: push eax
  loc_00DF129D: call [0040118Ch] ; __vbaI2Var
  loc_00DF12A3: xor ebx, ebx
  loc_00DF12A5: cmp ax, 0003h
  loc_00DF12A9: setz bl
  loc_00DF12AC: lea ecx, var_28
  loc_00DF12AF: neg ebx
  loc_00DF12B1: call [00401024h] ; __vbaFreeVar
  loc_00DF12B7: test bx, bx
  loc_00DF12BA: jz 00DF12E1h
  loc_00DF12BC: mov edx, [esi]
  loc_00DF12BE: lea eax, var_34
  loc_00DF12C1: lea ecx, var_30
  loc_00DF12C4: push eax
  loc_00DF12C5: push ecx
  loc_00DF12C6: push 00C0C0C0h
  loc_00DF12CB: push esi
  loc_00DF12CC: mov var_30, 00000000h
  loc_00DF12D3: call [edx+0000090Ch]
  loc_00DF12D9: mov edx, var_34
  loc_00DF12DC: mov var_14, edx
  loc_00DF12DF: jmp 00DF12EAh
  loc_00DF12E1: mov eax, [esi+000000A0h]
  loc_00DF12E7: mov var_14, eax
  loc_00DF12EA: cmp [edi], 00000000h
  loc_00DF12ED: jnz 00DF12FBh
  loc_00DF12EF: push edi
  loc_00DF12F0: push 006DEA0Ch
  loc_00DF12F5: call [004011A0h] ; __vbaNew2
  loc_00DF12FB: mov ecx, [edi]
  loc_00DF12FD: lea edx, var_18
  loc_00DF1300: push ecx
  loc_00DF1301: push edx
  loc_00DF1302: call [004010B4h] ; __vbaObjSetAddref
  loc_00DF1308: mov eax, [esi]
  loc_00DF130A: lea ecx, var_2C
  loc_00DF130D: lea edx, var_18
  loc_00DF1310: push ecx
  loc_00DF1311: push edx
  loc_00DF1312: push esi
  loc_00DF1313: call [eax+000009DCh]
  loc_00DF1319: mov eax, var_18
  loc_00DF131C: push 006DE764h
  loc_00DF1321: push eax
  loc_00DF1322: call [00401224h] ; __vbaCastObj
  loc_00DF1328: push eax
  loc_00DF1329: push edi
  loc_00DF132A: call [004010ACh] ; __vbaObjSet
  loc_00DF1330: mov bx, var_2C
  loc_00DF1334: lea ecx, var_18
  loc_00DF1337: call [00401254h] ; __vbaFreeObj
  loc_00DF133D: mov ecx, [esi]
  loc_00DF133F: lea edx, var_30
  loc_00DF1342: test bx, bx
  loc_00DF1345: push edx
  loc_00DF1346: push esi
  loc_00DF1347: jz 00DF13C6h
  loc_00DF1349: call [ecx+000000D8h]
  loc_00DF134F: test eax, eax
  loc_00DF1351: fnclex
  loc_00DF1353: jge 00DF1367h
  loc_00DF1355: push 000000D8h
  loc_00DF135A: push 006DCB00h
  loc_00DF135F: push esi
  loc_00DF1360: push eax
  loc_00DF1361: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF1367: cmp [edi], 00000000h
  loc_00DF136A: jnz 00DF1378h
  loc_00DF136C: push edi
  loc_00DF136D: push 006DEA0Ch
  loc_00DF1372: call [004011A0h] ; __vbaNew2
  loc_00DF1378: mov eax, arg_10
  loc_00DF137B: mov edx, [edi]
  loc_00DF137D: mov ebx, [esi]
  loc_00DF137F: push 00000000h
  loc_00DF1381: mov ecx, [eax]
  loc_00DF1383: lea eax, var_18
  loc_00DF1386: push ecx
  loc_00DF1387: push edx
  loc_00DF1388: push eax
  loc_00DF1389: call [004010B4h] ; __vbaObjSetAddref
  loc_00DF138F: mov ecx, [esi+00000174h]
  loc_00DF1395: mov edx, [esi+00000178h]
  loc_00DF139B: push eax
  loc_00DF139C: mov eax, arg_C
  loc_00DF139F: push ecx
  loc_00DF13A0: push edx
  loc_00DF13A1: mov ecx, [eax+00000004h]
  loc_00DF13A4: mov edx, [eax]
  loc_00DF13A6: mov eax, var_30
  loc_00DF13A9: push ecx
  loc_00DF13AA: push edx
  loc_00DF13AB: push eax
  loc_00DF13AC: push esi
  loc_00DF13AD: call [ebx+000008FCh]
  loc_00DF13B3: lea ecx, var_18
  loc_00DF13B6: call [00401254h] ; __vbaFreeObj
  loc_00DF13BC: push 00DF145Ah
  loc_00DF13C1: jmp 00DF1459h
  loc_00DF13C6: call [ecx+000000D8h]
  loc_00DF13CC: test eax, eax
  loc_00DF13CE: fnclex
  loc_00DF13D0: jge 00DF13E4h
  loc_00DF13D2: push 000000D8h
  loc_00DF13D7: push 006DCB00h
  loc_00DF13DC: push esi
  loc_00DF13DD: push eax
  loc_00DF13DE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF13E4: cmp [edi], 00000000h
  loc_00DF13E7: jnz 00DF13F5h
  loc_00DF13E9: push edi
  loc_00DF13EA: push 006DEA0Ch
  loc_00DF13EF: call [004011A0h] ; __vbaNew2
  loc_00DF13F5: mov eax, arg_10
  loc_00DF13F8: mov edx, var_14
  loc_00DF13FB: mov ebx, [esi]
  loc_00DF13FD: push 00000000h
  loc_00DF13FF: mov ecx, [eax]
  loc_00DF1401: mov eax, [edi]
  loc_00DF1403: push 00000000h
  loc_00DF1405: push ecx
  loc_00DF1406: push edx
  loc_00DF1407: lea ecx, var_18
  loc_00DF140A: push eax
  loc_00DF140B: push ecx
  loc_00DF140C: call [004010B4h] ; __vbaObjSetAddref
  loc_00DF1412: mov edx, [esi+00000174h]
  loc_00DF1418: push eax
  loc_00DF1419: mov eax, [esi+00000178h]
  loc_00DF141F: push edx
  loc_00DF1420: push eax
  loc_00DF1421: mov eax, arg_C
  loc_00DF1424: mov ecx, [eax+00000004h]
  loc_00DF1427: mov edx, [eax]
  loc_00DF1429: mov eax, var_30
  loc_00DF142C: push ecx
  loc_00DF142D: push edx
  loc_00DF142E: push eax
  loc_00DF142F: push esi
  loc_00DF1430: call [ebx+000008F8h]
  loc_00DF1436: lea ecx, var_18
  loc_00DF1439: call [00401254h] ; __vbaFreeObj
  loc_00DF143F: push 00DF145Ah
  loc_00DF1444: jmp 00DF1459h
  loc_00DF1446: lea ecx, var_18
  loc_00DF1449: call [00401254h] ; __vbaFreeObj
  loc_00DF144F: lea ecx, var_28
  loc_00DF1452: call [00401024h] ; __vbaFreeVar
  loc_00DF1458: ret
  loc_00DF1459: ret
  loc_00DF145A: mov ecx, var_10
  loc_00DF145D: pop edi
  loc_00DF145E: pop esi
  loc_00DF145F: xor eax, eax
  loc_00DF1461: mov fs:[00000000h], ecx
  loc_00DF1468: pop ebx
  loc_00DF1469: mov esp, ebp
  loc_00DF146B: pop ebp
  loc_00DF146C: retn 000Ch
End Sub

Private Sub Proc_2_118_DF1470(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20) 'DF1470
  loc_00DF1470: sub esp, 00000028h
  loc_00DF1473: xor eax, eax
  loc_00DF1475: push ebx
  loc_00DF1476: mov arg_14, eax
  loc_00DF147A: push esi
  loc_00DF147B: mov esi, arg_2C
  loc_00DF147F: mov arg_1C, eax
  loc_00DF1483: xor ebx, ebx
  loc_00DF1485: mov arg_20, eax
  loc_00DF1489: cmp [esi+00000170h], bx
  loc_00DF1490: push edi
  loc_00DF1491: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DF1497: mov arg_28, eax
  loc_00DF149B: mov arg_10, ebx
  loc_00DF149F: mov Proc_2_118_DF1470, ebx
  loc_00DF14A3: mov Me, ebx
  loc_00DF14A7: mov arg_C, ebx
  loc_00DF14AB: mov arg_14, ebx
  loc_00DF14AF: mov arg_18, ebx
  loc_00DF14B3: jz 00DF14CDh
  loc_00DF14B5: cmp [esi+00000048h], 00000001h
  loc_00DF14B9: jnz 00DF14CDh
  loc_00DF14BB: push FFFFFFFEh
  loc_00DF14BD: lea ecx, [esi+000001C0h]
  loc_00DF14C3: push FFFFFFFEh
  loc_00DF14C5: push ecx
  loc_00DF14C6: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF14CB: call edi
  loc_00DF14CD: mov edx, [esi]
  loc_00DF14CF: lea eax, Me
  loc_00DF14D3: lea ecx, Proc_2_118_DF1470
  loc_00DF14D7: push eax
  loc_00DF14D8: push ecx
  loc_00DF14D9: push 00808080h
  loc_00DF14DE: push esi
  loc_00DF14DF: mov arg_14, ebx
  loc_00DF14E3: call [edx+0000090Ch]
  loc_00DF14E9: mov edx, [esi]
  loc_00DF14EB: lea eax, arg_14
  loc_00DF14EF: push eax
  loc_00DF14F0: mov eax, [esi+00000088h]
  loc_00DF14F6: lea ecx, arg_10
  loc_00DF14FA: mov arg_10, ebx
  loc_00DF14FE: push ecx
  loc_00DF14FF: push eax
  loc_00DF1500: push esi
  loc_00DF1501: call [edx+0000090Ch]
  loc_00DF1507: mov ecx, [esi]
  loc_00DF1509: lea edx, arg_18
  loc_00DF150D: mov eax, arg_14
  loc_00DF1511: push edx
  loc_00DF1512: mov edx, arg_C
  loc_00DF1516: push eax
  loc_00DF1517: push edx
  loc_00DF1518: push esi
  loc_00DF1519: call [ecx+000008ECh]
  loc_00DF151F: lea ebx, [esi+000001C0h]
  loc_00DF1525: lea ecx, arg_1C
  loc_00DF1529: mov eax, arg_18
  loc_00DF152D: push ebx
  loc_00DF152E: push ecx
  loc_00DF152F: mov arg_18, eax
  loc_00DF1533: call 006DD98Ch ; CopyRect(%x1v, %x2v)
  loc_00DF1538: call edi
  loc_00DF153A: push 00000002h
  loc_00DF153C: lea edx, arg_20
  loc_00DF1540: push 00000002h
  loc_00DF1542: push edx
  loc_00DF1543: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF1548: call edi
  loc_00DF154A: mov eax, [esi]
  loc_00DF154C: lea ecx, Me
  loc_00DF1550: push ecx
  loc_00DF1551: lea edx, Me
  loc_00DF1555: lea ecx, arg_14
  loc_00DF1559: push edx
  loc_00DF155A: push ecx
  loc_00DF155B: push esi
  loc_00DF155C: mov arg_14, 3D4CCCCDh
  loc_00DF1564: call [eax+0000097Ch]
  loc_00DF156A: mov eax, [esi]
  loc_00DF156C: lea ecx, arg_C
  loc_00DF1570: mov edx, Me
  loc_00DF1574: push ecx
  loc_00DF1575: mov arg_10, edx
  loc_00DF1579: lea edx, arg_20
  loc_00DF157D: push edx
  loc_00DF157E: push esi
  loc_00DF157F: call [eax+00000928h]
  loc_00DF1585: push FFFFFFFFh
  loc_00DF1587: lea eax, arg_20
  loc_00DF158B: push FFFFFFFFh
  loc_00DF158D: push eax
  loc_00DF158E: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF1593: call edi
  loc_00DF1595: fld real8 ptr [004013B0h]
  loc_00DF159B: fstp real4 ptr Proc_2_118_DF1470
  loc_00DF159F: fnstsw ax
  loc_00DF15A1: test al, 0Dh
  loc_00DF15A3: jnz 00DF15FAh
  loc_00DF15A5: mov ecx, [esi]
  loc_00DF15A7: lea edx, Me
  loc_00DF15AB: push edx
  loc_00DF15AC: lea eax, Me
  loc_00DF15B0: lea edx, arg_14
  loc_00DF15B4: push eax
  loc_00DF15B5: push edx
  loc_00DF15B6: push esi
  loc_00DF15B7: call [ecx+0000097Ch]
  loc_00DF15BD: mov eax, Me
  loc_00DF15C1: mov ecx, [esi]
  loc_00DF15C3: mov arg_C, eax
  loc_00DF15C7: lea edx, arg_C
  loc_00DF15CB: lea eax, arg_1C
  loc_00DF15CF: push edx
  loc_00DF15D0: push eax
  loc_00DF15D1: push esi
  loc_00DF15D2: call [ecx+00000928h]
  loc_00DF15D8: mov ecx, [esi]
  loc_00DF15DA: lea edx, Proc_2_118_DF1470
  loc_00DF15DE: push edx
  loc_00DF15DF: push ebx
  loc_00DF15E0: push esi
  loc_00DF15E1: mov arg_10, FFFFFFFFh
  loc_00DF15E9: call [ecx+00000928h]
  loc_00DF15EF: pop edi
  loc_00DF15F0: pop esi
  loc_00DF15F1: xor eax, eax
  loc_00DF15F3: pop ebx
  loc_00DF15F4: add esp, 00000028h
  loc_00DF15F7: retn 0004h
End Sub

Private Sub Proc_2_119_DF1600(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28, arg_2C) 'DF1600
  loc_00DF1600: sub esp, 00000018h
  loc_00DF1603: push ebx
  loc_00DF1604: push ebp
  loc_00DF1605: push esi
  loc_00DF1606: mov esi, arg_20
  loc_00DF160A: lea ecx, Proc_2_119_DF1600
  loc_00DF160E: push edi
  loc_00DF160F: mov eax, [esi]
  loc_00DF1611: push ecx
  loc_00DF1612: mov ecx, [esi+00000088h]
  loc_00DF1618: lea edx, arg_10
  loc_00DF161C: push edx
  loc_00DF161D: xor edi, edi
  loc_00DF161F: push ecx
  loc_00DF1620: push esi
  loc_00DF1621: mov arg_2C, edi
  loc_00DF1625: mov arg_18, edi
  loc_00DF1629: mov arg_24, edi
  loc_00DF162D: mov arg_20, edi
  loc_00DF1631: mov arg_28, edi
  loc_00DF1635: mov arg_1C, edi
  loc_00DF1639: call [eax+0000090Ch]
  loc_00DF163F: mov eax, [esi+00000068h]
  loc_00DF1642: mov edx, Me
  loc_00DF1646: dec eax
  loc_00DF1647: cmp eax, 00000004h
  loc_00DF164A: mov arg_1C, edx
  loc_00DF164E: ja 00DF19CEh
  loc_00DF1654: jmp [eax*4+00DF1A7Ch]
  loc_00DF165B: mov eax, [esi]
  loc_00DF165D: lea ecx, Me
  loc_00DF1661: push ecx
  loc_00DF1662: lea edx, arg_10
  loc_00DF1666: lea ecx, arg_20
  loc_00DF166A: push edx
  loc_00DF166B: push ecx
  loc_00DF166C: push esi
  loc_00DF166D: mov arg_1C, 3E0F5C29h
  loc_00DF1675: call [eax+0000097Ch]
  loc_00DF167B: or ebp, FFFFFFFFh
  loc_00DF167E: mov arg_18, ebp
  loc_00DF1682: mov arg_10, ebp
  loc_00DF1686: jmp 00DF16B5h
  loc_00DF1688: mov eax, [esi]
  loc_00DF168A: lea ecx, Me
  loc_00DF168E: push ecx
  loc_00DF168F: lea edx, arg_10
  loc_00DF1693: lea ecx, arg_20
  loc_00DF1697: push edx
  loc_00DF1698: push ecx
  loc_00DF1699: push esi
  loc_00DF169A: mov arg_1C, 3E0F5C29h
  loc_00DF16A2: call [eax+0000097Ch]
  loc_00DF16A8: mov ebx, 00000001h
  loc_00DF16AD: mov arg_18, ebx
  loc_00DF16B1: mov arg_10, ebx
  loc_00DF16B5: mov edx, Me
  loc_00DF16B9: mov eax, [esi]
  loc_00DF16BB: mov arg_14, edx
  loc_00DF16BF: lea ecx, arg_18
  loc_00DF16C3: lea edx, arg_10
  loc_00DF16C7: push ecx
  loc_00DF16C8: push edx
  loc_00DF16C9: lea ecx, arg_1C
  loc_00DF16CD: lea edx, [esi+0000013Ch]
  loc_00DF16D3: push ecx
  loc_00DF16D4: push edx
  loc_00DF16D5: push esi
  loc_00DF16D6: call [eax+00000934h]
  loc_00DF16DC: jmp 00DF19CEh
  loc_00DF16E1: mov eax, [esi]
  loc_00DF16E3: lea ecx, Me
  loc_00DF16E7: lea edx, arg_C
  loc_00DF16EB: push ecx
  loc_00DF16EC: push edx
  loc_00DF16ED: push 00C0C0C0h
  loc_00DF16F2: push esi
  loc_00DF16F3: mov arg_1C, edi
  loc_00DF16F7: call [eax+0000090Ch]
  loc_00DF16FD: mov ecx, [esi]
  loc_00DF16FF: lea edx, arg_18
  loc_00DF1703: mov eax, Me
  loc_00DF1707: push edx
  loc_00DF1708: mov arg_18, eax
  loc_00DF170C: lea eax, arg_14
  loc_00DF1710: push eax
  loc_00DF1711: lea edx, arg_1C
  loc_00DF1715: lea eax, [esi+0000013Ch]
  loc_00DF171B: push edx
  loc_00DF171C: mov ebx, 00000001h
  loc_00DF1721: push eax
  loc_00DF1722: push esi
  loc_00DF1723: mov arg_2C, ebx
  loc_00DF1727: mov arg_24, ebx
  loc_00DF172B: call [ecx+00000934h]
  loc_00DF1731: jmp 00DF19CEh
  loc_00DF1736: mov ecx, [esi]
  loc_00DF1738: lea edx, Me
  loc_00DF173C: push edx
  loc_00DF173D: lea eax, arg_10
  loc_00DF1741: lea edx, arg_20
  loc_00DF1745: push eax
  loc_00DF1746: mov ebp, 3DCCCCCDh
  loc_00DF174B: push edx
  loc_00DF174C: push esi
  loc_00DF174D: mov arg_1C, ebp
  loc_00DF1751: call [ecx+0000097Ch]
  loc_00DF1757: mov ecx, [esi]
  loc_00DF1759: lea edx, arg_18
  loc_00DF175D: mov eax, Me
  loc_00DF1761: push edx
  loc_00DF1762: mov arg_18, eax
  loc_00DF1766: lea eax, arg_14
  loc_00DF176A: lea edx, arg_18
  loc_00DF176E: lea edi, [esi+0000013Ch]
  loc_00DF1774: push eax
  loc_00DF1775: push edx
  loc_00DF1776: mov ebx, 00000001h
  loc_00DF177B: push edi
  loc_00DF177C: push esi
  loc_00DF177D: mov arg_2C, ebx
  loc_00DF1781: mov arg_24, ebx
  loc_00DF1785: call [ecx+00000934h]
  loc_00DF178B: mov eax, [esi]
  loc_00DF178D: lea ecx, Me
  loc_00DF1791: push ecx
  loc_00DF1792: lea edx, arg_10
  loc_00DF1796: lea ecx, arg_20
  loc_00DF179A: push edx
  loc_00DF179B: push ecx
  loc_00DF179C: push esi
  loc_00DF179D: mov arg_1C, ebp
  loc_00DF17A1: call [eax+0000097Ch]
  loc_00DF17A7: mov eax, [esi]
  loc_00DF17A9: lea ecx, arg_18
  loc_00DF17AD: mov edx, Me
  loc_00DF17B1: push ecx
  loc_00DF17B2: mov arg_18, edx
  loc_00DF17B6: lea edx, arg_14
  loc_00DF17BA: lea ecx, arg_18
  loc_00DF17BE: push edx
  loc_00DF17BF: push ecx
  loc_00DF17C0: or ebp, FFFFFFFFh
  loc_00DF17C3: push edi
  loc_00DF17C4: push esi
  loc_00DF17C5: mov arg_2C, ebp
  loc_00DF17C9: mov arg_24, ebx
  loc_00DF17CD: call [eax+00000934h]
  loc_00DF17D3: mov edx, [esi]
  loc_00DF17D5: lea eax, Me
  loc_00DF17D9: push eax
  loc_00DF17DA: lea ecx, arg_10
  loc_00DF17DE: lea eax, arg_20
  loc_00DF17E2: push ecx
  loc_00DF17E3: push eax
  loc_00DF17E4: push esi
  loc_00DF17E5: mov arg_1C, 3DCCCCCDh
  loc_00DF17ED: call [edx+0000097Ch]
  loc_00DF17F3: mov edx, [esi]
  loc_00DF17F5: lea eax, arg_18
  loc_00DF17F9: mov ecx, Me
  loc_00DF17FD: push eax
  loc_00DF17FE: mov arg_18, ecx
  loc_00DF1802: lea ecx, arg_14
  loc_00DF1806: lea eax, arg_18
  loc_00DF180A: push ecx
  loc_00DF180B: push eax
  loc_00DF180C: push edi
  loc_00DF180D: push esi
  loc_00DF180E: mov arg_2C, ebx
  loc_00DF1812: mov arg_24, ebp
  loc_00DF1816: call [edx+00000934h]
  loc_00DF181C: mov ecx, [esi]
  loc_00DF181E: lea edx, Me
  loc_00DF1822: mov arg_C, 3DCCCCCDh
  loc_00DF182A: push edx
  loc_00DF182B: lea eax, arg_10
  loc_00DF182F: lea edx, arg_20
  loc_00DF1833: push eax
  loc_00DF1834: push edx
  loc_00DF1835: push esi
  loc_00DF1836: call [ecx+0000097Ch]
  loc_00DF183C: mov ecx, [esi]
  loc_00DF183E: lea edx, arg_18
  loc_00DF1842: mov eax, Me
  loc_00DF1846: push edx
  loc_00DF1847: mov arg_18, eax
  loc_00DF184B: lea eax, arg_14
  loc_00DF184F: lea edx, arg_18
  loc_00DF1853: push eax
  loc_00DF1854: push edx
  loc_00DF1855: push edi
  loc_00DF1856: push esi
  loc_00DF1857: mov arg_2C, ebp
  loc_00DF185B: mov arg_24, ebp
  loc_00DF185F: call [ecx+00000934h]
  loc_00DF1865: jmp 00DF19CCh
  loc_00DF186A: fld real8 ptr [004013B0h]
  loc_00DF1870: fstp real4 ptr arg_C
  loc_00DF1874: fnstsw ax
  loc_00DF1876: test al, 0Dh
  loc_00DF1878: jnz 00DF1A90h
  loc_00DF187E: lea ecx, Me
  loc_00DF1882: lea edx, arg_C
  loc_00DF1886: push ecx
  loc_00DF1887: lea ecx, arg_20
  loc_00DF188B: mov eax, [esi]
  loc_00DF188D: push edx
  loc_00DF188E: push ecx
  loc_00DF188F: push esi
  loc_00DF1890: call [eax+0000097Ch]
  loc_00DF1896: mov eax, [esi]
  loc_00DF1898: lea ecx, arg_18
  loc_00DF189C: mov edx, Me
  loc_00DF18A0: push ecx
  loc_00DF18A1: mov arg_18, edx
  loc_00DF18A5: lea edx, arg_14
  loc_00DF18A9: lea ecx, arg_18
  loc_00DF18AD: lea edi, [esi+0000013Ch]
  loc_00DF18B3: push edx
  loc_00DF18B4: push ecx
  loc_00DF18B5: mov ebx, 00000001h
  loc_00DF18BA: push edi
  loc_00DF18BB: push esi
  loc_00DF18BC: mov arg_2C, ebx
  loc_00DF18C0: mov arg_24, ebx
  loc_00DF18C4: call [eax+00000934h]
  loc_00DF18CA: fld real8 ptr [004013B0h]
  loc_00DF18D0: fstp real4 ptr arg_C
  loc_00DF18D4: fnstsw ax
  loc_00DF18D6: test al, 0Dh
  loc_00DF18D8: jnz 00DF1A90h
  loc_00DF18DE: mov edx, [esi]
  loc_00DF18E0: lea ecx, arg_C
  loc_00DF18E4: lea eax, Me
  loc_00DF18E8: push eax
  loc_00DF18E9: lea eax, arg_20
  loc_00DF18ED: push ecx
  loc_00DF18EE: push eax
  loc_00DF18EF: push esi
  loc_00DF18F0: call [edx+0000097Ch]
  loc_00DF18F6: mov edx, [esi]
  loc_00DF18F8: lea eax, arg_18
  loc_00DF18FC: mov ecx, Me
  loc_00DF1900: push eax
  loc_00DF1901: mov arg_18, ecx
  loc_00DF1905: lea ecx, arg_14
  loc_00DF1909: lea eax, arg_18
  loc_00DF190D: push ecx
  loc_00DF190E: push eax
  loc_00DF190F: or ebp, FFFFFFFFh
  loc_00DF1912: push edi
  loc_00DF1913: push esi
  loc_00DF1914: mov arg_2C, ebp
  loc_00DF1918: mov arg_24, ebx
  loc_00DF191C: call [edx+00000934h]
  loc_00DF1922: fld real8 ptr [004013B0h]
  loc_00DF1928: fstp real4 ptr arg_C
  loc_00DF192C: fnstsw ax
  loc_00DF192E: test al, 0Dh
  loc_00DF1930: jnz 00DF1A90h
  loc_00DF1936: mov ecx, [esi]
  loc_00DF1938: lea edx, Me
  loc_00DF193C: push edx
  loc_00DF193D: lea eax, arg_10
  loc_00DF1941: lea edx, arg_20
  loc_00DF1945: push eax
  loc_00DF1946: push edx
  loc_00DF1947: push esi
  loc_00DF1948: call [ecx+0000097Ch]
  loc_00DF194E: mov ecx, [esi]
  loc_00DF1950: lea edx, arg_18
  loc_00DF1954: mov eax, Me
  loc_00DF1958: push edx
  loc_00DF1959: mov arg_18, eax
  loc_00DF195D: lea eax, arg_14
  loc_00DF1961: lea edx, arg_18
  loc_00DF1965: push eax
  loc_00DF1966: mov arg_20, ebx
  loc_00DF196A: mov arg_18, ebp
  loc_00DF196E: push edx
  loc_00DF196F: push edi
  loc_00DF1970: push esi
  loc_00DF1971: call [ecx+00000934h]
  loc_00DF1977: fld real8 ptr [004013B0h]
  loc_00DF197D: fstp real4 ptr arg_C
  loc_00DF1981: fnstsw ax
  loc_00DF1983: test al, 0Dh
  loc_00DF1985: jnz 00DF1A90h
  loc_00DF198B: lea ecx, Me
  loc_00DF198F: lea edx, arg_C
  loc_00DF1993: push ecx
  loc_00DF1994: lea ecx, arg_20
  loc_00DF1998: mov eax, [esi]
  loc_00DF199A: push edx
  loc_00DF199B: push ecx
  loc_00DF199C: push esi
  loc_00DF199D: call [eax+0000097Ch]
  loc_00DF19A3: mov eax, [esi]
  loc_00DF19A5: lea ecx, arg_18
  loc_00DF19A9: mov edx, Me
  loc_00DF19AD: push ecx
  loc_00DF19AE: mov arg_18, edx
  loc_00DF19B2: lea edx, arg_14
  loc_00DF19B6: lea ecx, arg_18
  loc_00DF19BA: push edx
  loc_00DF19BB: push ecx
  loc_00DF19BC: push edi
  loc_00DF19BD: push esi
  loc_00DF19BE: mov arg_2C, ebp
  loc_00DF19C2: mov arg_24, ebp
  loc_00DF19C6: call [eax+00000934h]
  loc_00DF19CC: xor edi, edi
  loc_00DF19CE: cmp [esi+0000007Ch], di
  loc_00DF19D2: jz 00DF1A2Dh
  loc_00DF19D4: mov edx, [esi]
  loc_00DF19D6: lea eax, Me
  loc_00DF19DA: push eax
  loc_00DF19DB: mov eax, [esi+00000090h]
  loc_00DF19E1: lea ecx, arg_10
  loc_00DF19E5: mov arg_10, edi
  loc_00DF19E9: push ecx
  loc_00DF19EA: push eax
  loc_00DF19EB: push esi
  loc_00DF19EC: call [edx+0000090Ch]
  loc_00DF19F2: mov edx, [esi]
  loc_00DF19F4: lea eax, arg_18
  loc_00DF19F8: mov ecx, Me
  loc_00DF19FC: push eax
  loc_00DF19FD: mov arg_18, ecx
  loc_00DF1A01: lea ecx, arg_14
  loc_00DF1A05: push ecx
  loc_00DF1A06: lea eax, arg_1C
  loc_00DF1A0A: lea ecx, [esi+0000013Ch]
  loc_00DF1A10: push eax
  loc_00DF1A11: push ecx
  loc_00DF1A12: push esi
  loc_00DF1A13: mov arg_2C, edi
  loc_00DF1A17: mov arg_24, edi
  loc_00DF1A1B: call [edx+00000934h]
  loc_00DF1A21: pop edi
  loc_00DF1A22: pop esi
  loc_00DF1A23: pop ebp
  loc_00DF1A24: xor eax, eax
  loc_00DF1A26: pop ebx
  loc_00DF1A27: add esp, 00000018h
  loc_00DF1A2A: retn 0004h
End Sub

Private Function Proc_2_120_DF1AA0(arg_C, arg_10, arg_14, arg_18) 'DF1AA0
  loc_00DF1AA0: push ebp
  loc_00DF1AA1: mov ebp, esp
  loc_00DF1AA3: sub esp, 00000008h
  loc_00DF1AA6: push 00402836h ; __vbaExceptHandler
  loc_00DF1AAB: mov eax, fs:[00000000h]
  loc_00DF1AB1: push eax
  loc_00DF1AB2: mov fs:[00000000h], esp
  loc_00DF1AB9: sub esp, 000000A8h
  loc_00DF1ABF: push ebx
  loc_00DF1AC0: push esi
  loc_00DF1AC1: push edi
  loc_00DF1AC2: mov var_8, esp
  loc_00DF1AC5: mov var_4, 00401408h
  loc_00DF1ACC: mov esi, Me
  loc_00DF1ACF: xor eax, eax
  loc_00DF1AD1: mov var_20, eax
  loc_00DF1AD4: lea edx, var_AC
  loc_00DF1ADA: mov ecx, [esi]
  loc_00DF1ADC: mov var_1C, eax
  loc_00DF1ADF: mov var_18, eax
  loc_00DF1AE2: xor edi, edi
  loc_00DF1AE4: push edx
  loc_00DF1AE5: push esi
  loc_00DF1AE6: mov var_14, eax
  loc_00DF1AE9: mov var_28, edi
  loc_00DF1AEC: mov var_38, edi
  loc_00DF1AEF: mov var_48, edi
  loc_00DF1AF2: mov var_58, edi
  loc_00DF1AF5: mov var_68, edi
  loc_00DF1AF8: mov var_78, edi
  loc_00DF1AFB: mov var_A8, edi
  loc_00DF1B01: mov var_AC, edi
  loc_00DF1B07: call [ecx+000000D8h]
  loc_00DF1B0D: cmp eax, edi
  loc_00DF1B0F: fnclex
  loc_00DF1B11: jge 00DF1B29h
  loc_00DF1B13: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF1B19: push 000000D8h
  loc_00DF1B1E: push 006DCB00h
  loc_00DF1B23: push esi
  loc_00DF1B24: push eax
  loc_00DF1B25: call ebx
  loc_00DF1B27: jmp 00DF1B2Fh
  loc_00DF1B29: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF1B2F: mov eax, var_AC
  loc_00DF1B35: push eax
  loc_00DF1B36: call 006DD71Ch ; GetTextColor(%x1v)
  loc_00DF1B3B: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DF1B41: mov var_B0, eax
  loc_00DF1B47: call edi
  loc_00DF1B49: mov edx, arg_C
  loc_00DF1B4C: mov ecx, var_B0
  loc_00DF1B52: lea eax, var_20
  loc_00DF1B55: push edx
  loc_00DF1B56: push eax
  loc_00DF1B57: mov var_24, ecx
  loc_00DF1B5A: call 006DD98Ch ; CopyRect(%x1v, %x2v)
  loc_00DF1B5F: call edi
  loc_00DF1B61: mov ecx, arg_18
  loc_00DF1B64: mov eax, arg_14
  loc_00DF1B67: mov edx, [ecx]
  loc_00DF1B69: mov ecx, [eax]
  loc_00DF1B6B: push edx
  loc_00DF1B6C: lea edx, var_20
  loc_00DF1B6F: push ecx
  loc_00DF1B70: push edx
  loc_00DF1B71: call 006DD948h ; OffsetRect(%x1v, %x2v, %x3v)
  loc_00DF1B76: call edi
  loc_00DF1B78: mov eax, [esi]
  loc_00DF1B7A: lea ecx, var_AC
  loc_00DF1B80: push ecx
  loc_00DF1B81: push esi
  loc_00DF1B82: call [eax+000000D8h]
  loc_00DF1B88: test eax, eax
  loc_00DF1B8A: fnclex
  loc_00DF1B8C: jge 00DF1B9Ch
  loc_00DF1B8E: push 000000D8h
  loc_00DF1B93: push 006DCB00h
  loc_00DF1B98: push esi
  loc_00DF1B99: push eax
  loc_00DF1B9A: call ebx
  loc_00DF1B9C: mov edx, arg_10
  loc_00DF1B9F: mov ecx, var_AC
  loc_00DF1BA5: mov eax, [edx]
  loc_00DF1BA7: push eax
  loc_00DF1BA8: push ecx
  loc_00DF1BA9: call 006DD780h ; SetTextColor(%x1v, %x2v)
  loc_00DF1BAE: call edi
  loc_00DF1BB0: mov eax, [esi+00000078h]
  loc_00DF1BB3: test eax, eax
  loc_00DF1BB5: jz 00DF1C99h
  loc_00DF1BBB: mov edx, [esi]
  loc_00DF1BBD: lea eax, var_AC
  loc_00DF1BC3: push eax
  loc_00DF1BC4: push esi
  loc_00DF1BC5: call [edx+000000D8h]
  loc_00DF1BCB: test eax, eax
  loc_00DF1BCD: fnclex
  loc_00DF1BCF: jge 00DF1BDFh
  loc_00DF1BD1: push 000000D8h
  loc_00DF1BD6: push 006DCB00h
  loc_00DF1BDB: push esi
  loc_00DF1BDC: push eax
  loc_00DF1BDD: call ebx
  loc_00DF1BDF: mov ecx, [esi+00000080h]
  loc_00DF1BE5: push ecx
  loc_00DF1BE6: call [00401190h] ; VarPtr
  loc_00DF1BEC: mov var_B0, eax
  loc_00DF1BF2: mov eax, 00000003h
  loc_00DF1BF7: mov var_A8, eax
  loc_00DF1BFD: mov var_38, eax
  loc_00DF1C00: lea edx, [esi+00000138h]
  loc_00DF1C06: lea eax, var_48
  loc_00DF1C09: mov var_70, edx
  loc_00DF1C0C: lea ecx, var_38
  loc_00DF1C0F: push eax
  loc_00DF1C10: lea edx, var_78
  loc_00DF1C13: push ecx
  loc_00DF1C14: lea eax, var_58
  loc_00DF1C17: push edx
  loc_00DF1C18: push eax
  loc_00DF1C19: mov var_A0, 00000011h
  loc_00DF1C23: mov var_40, 00000000h
  loc_00DF1C2A: mov var_48, 00000002h
  loc_00DF1C31: mov var_30, 00020000h
  loc_00DF1C38: mov var_78, 0000400Bh
  loc_00DF1C3F: call [004011B4h] ; rtcImmediateIf
  loc_00DF1C45: lea ecx, var_A8
  loc_00DF1C4B: lea edx, var_58
  loc_00DF1C4E: push ecx
  loc_00DF1C4F: lea eax, var_68
  loc_00DF1C52: push edx
  loc_00DF1C53: push eax
  loc_00DF1C54: call [00401118h] ; __vbaVarOr
  loc_00DF1C5A: push eax
  loc_00DF1C5B: call [004011D0h] ; __vbaI4Var
  loc_00DF1C61: mov edx, var_B0
  loc_00DF1C67: lea ecx, var_20
  loc_00DF1C6A: push eax
  loc_00DF1C6B: mov eax, var_AC
  loc_00DF1C71: push ecx
  loc_00DF1C72: push FFFFFFFFh
  loc_00DF1C74: push edx
  loc_00DF1C75: push eax
  loc_00DF1C76: call 006DE300h ; DrawTextW()
  loc_00DF1C7B: call edi
  loc_00DF1C7D: lea ecx, var_58
  loc_00DF1C80: lea edx, var_48
  loc_00DF1C83: push ecx
  loc_00DF1C84: lea eax, var_38
  loc_00DF1C87: push edx
  loc_00DF1C88: push eax
  loc_00DF1C89: push 00000003h
  loc_00DF1C8B: call [00401038h] ; __vbaFreeVarList
  loc_00DF1C91: add esp, 00000010h
  loc_00DF1C94: jmp 00DF1D86h
  loc_00DF1C99: mov ecx, [esi]
  loc_00DF1C9B: lea edx, var_AC
  loc_00DF1CA1: push edx
  loc_00DF1CA2: push esi
  loc_00DF1CA3: call [ecx+000000D8h]
  loc_00DF1CA9: test eax, eax
  loc_00DF1CAB: fnclex
  loc_00DF1CAD: jge 00DF1CBDh
  loc_00DF1CAF: push 000000D8h
  loc_00DF1CB4: push 006DCB00h
  loc_00DF1CB9: push esi
  loc_00DF1CBA: push eax
  loc_00DF1CBB: call ebx
  loc_00DF1CBD: mov eax, 00000003h
  loc_00DF1CC2: lea ecx, var_48
  loc_00DF1CC5: mov var_A8, eax
  loc_00DF1CCB: mov var_38, eax
  loc_00DF1CCE: lea eax, [esi+00000138h]
  loc_00DF1CD4: lea edx, var_38
  loc_00DF1CD7: mov var_70, eax
  loc_00DF1CDA: push ecx
  loc_00DF1CDB: lea eax, var_78
  loc_00DF1CDE: push edx
  loc_00DF1CDF: lea ecx, var_58
  loc_00DF1CE2: push eax
  loc_00DF1CE3: push ecx
  loc_00DF1CE4: mov var_A0, 00000011h
  loc_00DF1CEE: mov var_40, 00000000h
  loc_00DF1CF5: mov var_48, 00000002h
  loc_00DF1CFC: mov var_30, 00020000h
  loc_00DF1D03: mov var_78, 0000400Bh
  loc_00DF1D0A: call [004011B4h] ; rtcImmediateIf
  loc_00DF1D10: lea edx, var_A8
  loc_00DF1D16: lea eax, var_58
  loc_00DF1D19: push edx
  loc_00DF1D1A: lea ecx, var_68
  loc_00DF1D1D: push eax
  loc_00DF1D1E: push ecx
  loc_00DF1D1F: lea ebx, [esi+00000080h]
  loc_00DF1D25: call [00401118h] ; __vbaVarOr
  loc_00DF1D2B: push eax
  loc_00DF1D2C: call [004011D0h] ; __vbaI4Var
  loc_00DF1D32: lea edx, var_20
  loc_00DF1D35: push eax
  loc_00DF1D36: mov eax, [ebx]
  loc_00DF1D38: push edx
  loc_00DF1D39: push FFFFFFFFh
  loc_00DF1D3B: lea ecx, var_28
  loc_00DF1D3E: push eax
  loc_00DF1D3F: push ecx
  loc_00DF1D40: call [004011ECh] ; __vbaStrToAnsi
  loc_00DF1D46: mov edx, var_AC
  loc_00DF1D4C: push eax
  loc_00DF1D4D: push edx
  loc_00DF1D4E: call 006DE2BCh ; DrawText(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF1D53: call edi
  loc_00DF1D55: mov eax, var_28
  loc_00DF1D58: push eax
  loc_00DF1D59: push ebx
  loc_00DF1D5A: call [00401160h] ; __vbaStrToUnicode
  loc_00DF1D60: lea ecx, var_28
  loc_00DF1D63: call [00401258h] ; __vbaFreeStr
  loc_00DF1D69: lea ecx, var_58
  loc_00DF1D6C: lea edx, var_48
  loc_00DF1D6F: push ecx
  loc_00DF1D70: lea eax, var_38
  loc_00DF1D73: push edx
  loc_00DF1D74: push eax
  loc_00DF1D75: push 00000003h
  loc_00DF1D77: call [00401038h] ; __vbaFreeVarList
  loc_00DF1D7D: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF1D83: add esp, 00000010h
  loc_00DF1D86: mov ecx, [esi]
  loc_00DF1D88: lea edx, var_AC
  loc_00DF1D8E: push edx
  loc_00DF1D8F: push esi
  loc_00DF1D90: call [ecx+000000D8h]
  loc_00DF1D96: test eax, eax
  loc_00DF1D98: fnclex
  loc_00DF1D9A: jge 00DF1DAAh
  loc_00DF1D9C: push 000000D8h
  loc_00DF1DA1: push 006DCB00h
  loc_00DF1DA6: push esi
  loc_00DF1DA7: push eax
  loc_00DF1DA8: call ebx
  loc_00DF1DAA: mov eax, var_24
  loc_00DF1DAD: mov ecx, var_AC
  loc_00DF1DB3: push eax
  loc_00DF1DB4: push ecx
  loc_00DF1DB5: call 006DD780h ; SetTextColor(%x1v, %x2v)
  loc_00DF1DBA: call edi
  loc_00DF1DBC: push 00DF1DE9h
  loc_00DF1DC1: jmp 00DF1DE8h
  loc_00DF1DC3: lea ecx, var_28
  loc_00DF1DC6: call [00401258h] ; __vbaFreeStr
  loc_00DF1DCC: lea edx, var_68
  loc_00DF1DCF: lea eax, var_58
  loc_00DF1DD2: push edx
  loc_00DF1DD3: lea ecx, var_48
  loc_00DF1DD6: push eax
  loc_00DF1DD7: lea edx, var_38
  loc_00DF1DDA: push ecx
  loc_00DF1DDB: push edx
  loc_00DF1DDC: push 00000004h
  loc_00DF1DDE: call [00401038h] ; __vbaFreeVarList
  loc_00DF1DE4: add esp, 00000014h
  loc_00DF1DE7: ret
  loc_00DF1DE8: ret
  loc_00DF1DE9: mov ecx, var_10
  loc_00DF1DEC: pop edi
  loc_00DF1DED: pop esi
  loc_00DF1DEE: xor eax, eax
  loc_00DF1DF0: mov fs:[00000000h], ecx
  loc_00DF1DF7: pop ebx
  loc_00DF1DF8: mov esp, ebp
  loc_00DF1DFA: pop ebp
  loc_00DF1DFB: retn 0014h
End Function

Private Sub Proc_2_121_DF1E00
  loc_00DF1E00: push ebp
  loc_00DF1E01: mov ebp, esp
  loc_00DF1E03: sub esp, 00000008h
  loc_00DF1E06: push 00402836h ; __vbaExceptHandler
  loc_00DF1E0B: mov eax, fs:[00000000h]
  loc_00DF1E11: push eax
  loc_00DF1E12: mov fs:[00000000h], esp
  loc_00DF1E19: sub esp, 00000074h
  loc_00DF1E1C: push ebx
  loc_00DF1E1D: push esi
  loc_00DF1E1E: push edi
  loc_00DF1E1F: mov var_8, esp
  loc_00DF1E22: mov var_4, 00401418h
  loc_00DF1E29: mov edi, Me
  loc_00DF1E2C: lea ecx, var_18
  loc_00DF1E2F: xor esi, esi
  loc_00DF1E31: push ecx
  loc_00DF1E32: mov eax, [edi]
  loc_00DF1E34: push edi
  loc_00DF1E35: mov var_14, esi
  loc_00DF1E38: mov var_18, esi
  loc_00DF1E3B: mov var_28, esi
  loc_00DF1E3E: mov var_38, esi
  loc_00DF1E41: mov var_48, esi
  loc_00DF1E44: mov var_58, esi
  loc_00DF1E47: mov var_6C, esi
  loc_00DF1E4A: mov var_78, esi
  loc_00DF1E4D: mov var_7C, esi
  loc_00DF1E50: call [eax+00000198h]
  loc_00DF1E56: cmp eax, esi
  loc_00DF1E58: fnclex
  loc_00DF1E5A: jge 00DF1E6Eh
  loc_00DF1E5C: push 00000198h
  loc_00DF1E61: push 006DCB00h
  loc_00DF1E66: push edi
  loc_00DF1E67: push eax
  loc_00DF1E68: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF1E6E: mov edx, var_18
  loc_00DF1E71: mov edi, [00401214h] ; __vbaLateMemCallLd
  loc_00DF1E77: push esi
  loc_00DF1E78: push 006DEA1Ch ; "Controls"
  loc_00DF1E7D: lea eax, var_38
  loc_00DF1E80: push edx
  loc_00DF1E81: push eax
  loc_00DF1E82: call edi
  loc_00DF1E84: add esp, 00000010h
  loc_00DF1E87: push eax
  loc_00DF1E88: call [0040110Ch] ; __vbaObjVar
  loc_00DF1E8E: lea ecx, var_78
  loc_00DF1E91: push eax
  loc_00DF1E92: push ecx
  loc_00DF1E93: call [004010B4h] ; __vbaObjSetAddref
  loc_00DF1E99: push eax
  loc_00DF1E9A: lea edx, var_14
  loc_00DF1E9D: lea eax, var_7C
  loc_00DF1EA0: push edx
  loc_00DF1EA1: push eax
  loc_00DF1EA2: call [00401068h] ; __vbaForEachCollAd
  loc_00DF1EA8: lea ecx, var_18
  loc_00DF1EAB: mov esi, eax
  loc_00DF1EAD: call [00401254h] ; __vbaFreeObj
  loc_00DF1EB3: lea ecx, var_38
  loc_00DF1EB6: call [00401024h] ; __vbaFreeVar
  loc_00DF1EBC: mov ebx, [00401108h] ; __vbaVarTstEq
  loc_00DF1EC2: test esi, esi
  loc_00DF1EC4: jz 00DF206Ah
  loc_00DF1ECA: mov ecx, var_14
  loc_00DF1ECD: push 006DD090h
  loc_00DF1ED2: push ecx
  loc_00DF1ED3: call [00401188h] ; __vbaCheckType
  loc_00DF1ED9: test ax, ax
  loc_00DF1EDC: jz 00DF2055h
  loc_00DF1EE2: mov esi, Me
  loc_00DF1EE5: lea ecx, var_6C
  loc_00DF1EE8: push ecx
  loc_00DF1EE9: mov eax, [esi+00000010h]
  loc_00DF1EEC: push eax
  loc_00DF1EED: mov edx, [eax]
  loc_00DF1EEF: call [edx+00000268h]
  loc_00DF1EF5: test eax, eax
  loc_00DF1EF7: fnclex
  loc_00DF1EF9: jge 00DF1F10h
  loc_00DF1EFB: mov edx, [esi+00000010h]
  loc_00DF1EFE: push 00000268h
  loc_00DF1F03: push 006DCB00h
  loc_00DF1F08: push edx
  loc_00DF1F09: push eax
  loc_00DF1F0A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF1F10: mov ecx, var_14
  loc_00DF1F13: mov eax, var_6C
  loc_00DF1F16: push 00000000h
  loc_00DF1F18: push 006DEA44h ; "Hwnd"
  loc_00DF1F1D: push 00000000h
  loc_00DF1F1F: push 006DEA30h ; "Container"
  loc_00DF1F24: lea edx, var_28
  loc_00DF1F27: push ecx
  loc_00DF1F28: push edx
  loc_00DF1F29: mov var_50, eax
  loc_00DF1F2C: mov var_58, 00008003h
  loc_00DF1F33: call edi
  loc_00DF1F35: add esp, 00000010h
  loc_00DF1F38: push eax
  loc_00DF1F39: lea eax, var_38
  loc_00DF1F3C: push eax
  loc_00DF1F3D: call [0040120Ch] ; __vbaVarLateMemCallLd
  loc_00DF1F43: add esp, 00000010h
  loc_00DF1F46: lea ecx, var_58
  loc_00DF1F49: push eax
  loc_00DF1F4A: push ecx
  loc_00DF1F4B: call ebx
  loc_00DF1F4D: mov esi, eax
  loc_00DF1F4F: lea edx, var_38
  loc_00DF1F52: lea eax, var_28
  loc_00DF1F55: push edx
  loc_00DF1F56: push eax
  loc_00DF1F57: push 00000002h
  loc_00DF1F59: call [00401038h] ; __vbaFreeVarList
  loc_00DF1F5F: add esp, 0000000Ch
  loc_00DF1F62: test si, si
  loc_00DF1F65: jz 00DF2055h
  loc_00DF1F6B: mov ecx, var_14
  loc_00DF1F6E: push 00000000h
  loc_00DF1F70: push 006DEA50h ; "Mode"
  loc_00DF1F75: lea edx, var_28
  loc_00DF1F78: push ecx
  loc_00DF1F79: push edx
  loc_00DF1F7A: mov var_50, 00000002h
  loc_00DF1F81: mov var_58, 00008003h
  loc_00DF1F88: call edi
  loc_00DF1F8A: add esp, 00000010h
  loc_00DF1F8D: push eax
  loc_00DF1F8E: lea eax, var_58
  loc_00DF1F91: push eax
  loc_00DF1F92: call ebx
  loc_00DF1F94: lea ecx, var_28
  loc_00DF1F97: mov si, ax
  loc_00DF1F9A: call [00401024h] ; __vbaFreeVar
  loc_00DF1FA0: test si, si
  loc_00DF1FA3: jz 00DF2055h
  loc_00DF1FA9: mov esi, Me
  loc_00DF1FAC: lea edx, var_6C
  loc_00DF1FAF: push edx
  loc_00DF1FB0: mov eax, [esi+00000010h]
  loc_00DF1FB3: push eax
  loc_00DF1FB4: mov ecx, [eax]
  loc_00DF1FB6: call [ecx+00000058h]
  loc_00DF1FB9: test eax, eax
  loc_00DF1FBB: fnclex
  loc_00DF1FBD: jge 00DF1FD1h
  loc_00DF1FBF: mov ecx, [esi+00000010h]
  loc_00DF1FC2: push 00000058h
  loc_00DF1FC4: push 006DCB00h
  loc_00DF1FC9: push ecx
  loc_00DF1FCA: push eax
  loc_00DF1FCB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF1FD1: mov eax, var_14
  loc_00DF1FD4: mov edx, var_6C
  loc_00DF1FD7: push 00000000h
  loc_00DF1FD9: push 006DEA44h ; "Hwnd"
  loc_00DF1FDE: lea ecx, var_28
  loc_00DF1FE1: push eax
  loc_00DF1FE2: push ecx
  loc_00DF1FE3: mov var_50, edx
  loc_00DF1FE6: mov var_58, 00008003h
  loc_00DF1FED: call edi
  loc_00DF1FEF: add esp, 00000010h
  loc_00DF1FF2: lea edx, var_58
  loc_00DF1FF5: push eax
  loc_00DF1FF6: lea eax, var_38
  loc_00DF1FF9: push edx
  loc_00DF1FFA: push eax
  loc_00DF1FFB: call [004011D4h] ; __vbaVarCmpEq
  loc_00DF2001: lea ecx, var_48
  loc_00DF2004: push eax
  loc_00DF2005: push ecx
  loc_00DF2006: call [004011B8h] ; __vbaVarNot
  loc_00DF200C: push eax
  loc_00DF200D: call [004010DCh] ; __vbaBoolVarNull
  loc_00DF2013: lea ecx, var_28
  loc_00DF2016: mov esi, eax
  loc_00DF2018: call [00401024h] ; __vbaFreeVar
  loc_00DF201E: test si, si
  loc_00DF2021: jz 00DF2055h
  loc_00DF2023: sub esp, 00000010h
  loc_00DF2026: mov ecx, 0000000Bh
  loc_00DF202B: mov edx, esp
  loc_00DF202D: mov var_58, ecx
  loc_00DF2030: xor eax, eax
  loc_00DF2032: push 006DEA5Ch ; "Value"
  loc_00DF2037: mov [edx], ecx
  loc_00DF2039: mov ecx, var_54
  loc_00DF203C: mov var_50, eax
  loc_00DF203F: mov [edx+00000004h], ecx
  loc_00DF2042: mov ecx, var_14
  loc_00DF2045: push ecx
  loc_00DF2046: mov [edx+00000008h], eax
  loc_00DF2049: mov eax, var_4C
  loc_00DF204C: mov [edx+0000000Ch], eax
  loc_00DF204F: call [00401094h] ; __vbaLateMemSt
  loc_00DF2055: lea edx, var_14
  loc_00DF2058: lea eax, var_7C
  loc_00DF205B: push edx
  loc_00DF205C: push eax
  loc_00DF205D: call [00401240h] ; __vbaNextEachCollAd
  loc_00DF2063: mov esi, eax
  loc_00DF2065: jmp 00DF1EC2h
  loc_00DF206A: push 00DF20AFh
  loc_00DF206F: jmp 00DF2092h
  loc_00DF2071: lea ecx, var_18
  loc_00DF2074: call [00401254h] ; __vbaFreeObj
  loc_00DF207A: lea ecx, var_48
  loc_00DF207D: lea edx, var_38
  loc_00DF2080: push ecx
  loc_00DF2081: lea eax, var_28
  loc_00DF2084: push edx
  loc_00DF2085: push eax
  loc_00DF2086: push 00000003h
  loc_00DF2088: call [00401038h] ; __vbaFreeVarList
  loc_00DF208E: add esp, 00000010h
  loc_00DF2091: ret
  loc_00DF2092: lea ecx, var_7C
  loc_00DF2095: lea edx, var_78
  loc_00DF2098: push ecx
  loc_00DF2099: push edx
  loc_00DF209A: push 00000002h
  loc_00DF209C: call [00401048h] ; __vbaFreeObjList
  loc_00DF20A2: add esp, 0000000Ch
  loc_00DF20A5: lea ecx, var_14
  loc_00DF20A8: call [00401254h] ; __vbaFreeObj
  loc_00DF20AE: ret
  loc_00DF20AF: mov ecx, var_10
  loc_00DF20B2: pop edi
  loc_00DF20B3: pop esi
  loc_00DF20B4: xor eax, eax
  loc_00DF20B6: mov fs:[00000000h], ecx
  loc_00DF20BD: pop ebx
  loc_00DF20BE: mov esp, ebp
  loc_00DF20C0: pop ebp
  loc_00DF20C1: retn 0004h
End Sub

Private Sub Proc_2_122_DF20D0
  loc_00DF20D0: push ebp
  loc_00DF20D1: mov ebp, esp
  loc_00DF20D3: sub esp, 00000008h
  loc_00DF20D6: push 00402836h ; __vbaExceptHandler
  loc_00DF20DB: mov eax, fs:[00000000h]
  loc_00DF20E1: push eax
  loc_00DF20E2: mov fs:[00000000h], esp
  loc_00DF20E9: sub esp, 00000038h
  loc_00DF20EC: push ebx
  loc_00DF20ED: push esi
  loc_00DF20EE: push edi
  loc_00DF20EF: mov var_8, esp
  loc_00DF20F2: mov var_4, 00401428h
  loc_00DF20F9: mov edi, Me
  loc_00DF20FC: xor esi, esi
  loc_00DF20FE: mov var_18, esi
  loc_00DF2101: mov var_1C, esi
  loc_00DF2104: mov eax, [edi+00000010h]
  loc_00DF2107: mov var_2C, esi
  loc_00DF210A: push esi
  loc_00DF210B: push eax
  loc_00DF210C: mov ecx, [eax]
  loc_00DF210E: call [ecx+00000324h]
  loc_00DF2114: cmp eax, esi
  loc_00DF2116: fnclex
  loc_00DF2118: jge 00DF212Fh
  loc_00DF211A: mov edx, [edi+00000010h]
  loc_00DF211D: push 00000324h
  loc_00DF2122: push 006DCB00h
  loc_00DF2127: push edx
  loc_00DF2128: push eax
  loc_00DF2129: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF212F: mov eax, [edi+00000080h]
  loc_00DF2135: mov ebx, [0040102Ch] ; __vbaLenBstr
  loc_00DF213B: push eax
  loc_00DF213C: call ebx
  loc_00DF213E: cmp eax, 00000001h
  loc_00DF2141: jle 00DF2345h
  loc_00DF2147: mov ecx, [edi+00000080h]
  loc_00DF214D: push 00000001h
  loc_00DF214F: push ecx
  loc_00DF2150: push 006DEA6Ch ; "&"
  loc_00DF2155: push 00000001h
  loc_00DF2157: call [0040119Ch] ; __vbaInStr
  loc_00DF215D: mov edx, [edi+00000080h]
  loc_00DF2163: mov esi, eax
  loc_00DF2165: push edx
  loc_00DF2166: call ebx
  loc_00DF2168: cmp esi, eax
  loc_00DF216A: jge 00DF2345h
  loc_00DF2170: test esi, esi
  loc_00DF2172: jle 00DF2345h
  loc_00DF2178: mov edx, [edi+00000080h]
  loc_00DF217E: mov ecx, esi
  loc_00DF2180: lea eax, var_2C
  loc_00DF2183: add ecx, 00000001h
  loc_00DF2186: push eax
  loc_00DF2187: mov var_24, 00000001h
  loc_00DF218E: jo 00DF237Fh
  loc_00DF2194: push ecx
  loc_00DF2195: push edx
  loc_00DF2196: mov var_2C, 00000002h
  loc_00DF219D: call [004010ECh] ; rtcMidCharBstr
  loc_00DF21A3: mov edx, eax
  loc_00DF21A5: lea ecx, var_18
  loc_00DF21A8: call [00401228h] ; __vbaStrMove
  loc_00DF21AE: push eax
  loc_00DF21AF: push 006DEA6Ch ; "&"
  loc_00DF21B4: call [00401104h] ; __vbaStrCmp
  loc_00DF21BA: mov ebx, eax
  loc_00DF21BC: lea ecx, var_18
  loc_00DF21BF: neg ebx
  loc_00DF21C1: sbb ebx, ebx
  loc_00DF21C3: neg ebx
  loc_00DF21C5: neg ebx
  loc_00DF21C7: call [00401258h] ; __vbaFreeStr
  loc_00DF21CD: lea ecx, var_2C
  loc_00DF21D0: call [00401024h] ; __vbaFreeVar
  loc_00DF21D6: test bx, bx
  loc_00DF21D9: jz 00DF2238h
  loc_00DF21DB: mov ecx, [edi+00000080h]
  loc_00DF21E1: mov ebx, [edi]
  loc_00DF21E3: lea eax, var_2C
  loc_00DF21E6: add esi, 00000001h
  loc_00DF21E9: push eax
  loc_00DF21EA: mov var_24, 00000001h
  loc_00DF21F1: jo 00DF237Fh
  loc_00DF21F7: push esi
  loc_00DF21F8: push ecx
  loc_00DF21F9: mov var_2C, 00000002h
  loc_00DF2200: call [004010ECh] ; rtcMidCharBstr
  loc_00DF2206: mov esi, [00401228h] ; __vbaStrMove
  loc_00DF220C: mov edx, eax
  loc_00DF220E: lea ecx, var_18
  loc_00DF2211: call __vbaStrMove
  loc_00DF2213: push eax
  loc_00DF2214: call [0040104Ch] ; rtcLowerCaseBstr
  loc_00DF221A: mov edx, eax
  loc_00DF221C: lea ecx, var_1C
  loc_00DF221F: call __vbaStrMove
  loc_00DF2221: push eax
  loc_00DF2222: push edi
  loc_00DF2223: call [ebx+00000324h]
  loc_00DF2229: test eax, eax
  loc_00DF222B: fnclex
  loc_00DF222D: jge 00DF2329h
  loc_00DF2233: jmp 00DF2317h
  loc_00DF2238: mov ecx, [edi+00000080h]
  loc_00DF223E: add esi, 00000002h
  loc_00DF2241: jo 00DF237Fh
  loc_00DF2247: push esi
  loc_00DF2248: push ecx
  loc_00DF2249: push 006DEA6Ch ; "&"
  loc_00DF224E: push 00000001h
  loc_00DF2250: call [0040119Ch] ; __vbaInStr
  loc_00DF2256: mov ecx, [edi+00000080h]
  loc_00DF225C: mov ebx, eax
  loc_00DF225E: lea edx, var_2C
  loc_00DF2261: add eax, 00000001h
  loc_00DF2264: push edx
  loc_00DF2265: mov var_24, 00000001h
  loc_00DF226C: jo 00DF237Fh
  loc_00DF2272: push eax
  loc_00DF2273: push ecx
  loc_00DF2274: mov var_2C, 00000002h
  loc_00DF227B: call [004010ECh] ; rtcMidCharBstr
  loc_00DF2281: mov edx, eax
  loc_00DF2283: lea ecx, var_18
  loc_00DF2286: call [00401228h] ; __vbaStrMove
  loc_00DF228C: push eax
  loc_00DF228D: push 006DEA6Ch ; "&"
  loc_00DF2292: call [00401104h] ; __vbaStrCmp
  loc_00DF2298: mov esi, eax
  loc_00DF229A: lea ecx, var_18
  loc_00DF229D: neg esi
  loc_00DF229F: sbb esi, esi
  loc_00DF22A1: neg esi
  loc_00DF22A3: neg esi
  loc_00DF22A5: call [00401258h] ; __vbaFreeStr
  loc_00DF22AB: lea ecx, var_2C
  loc_00DF22AE: call [00401024h] ; __vbaFreeVar
  loc_00DF22B4: test si, si
  loc_00DF22B7: jz 00DF2345h
  loc_00DF22BD: mov eax, [edi+00000080h]
  loc_00DF22C3: mov esi, [edi]
  loc_00DF22C5: lea edx, var_2C
  loc_00DF22C8: add ebx, 00000001h
  loc_00DF22CB: push edx
  loc_00DF22CC: mov var_24, 00000001h
  loc_00DF22D3: jo 00DF237Fh
  loc_00DF22D9: push ebx
  loc_00DF22DA: push eax
  loc_00DF22DB: mov var_2C, 00000002h
  loc_00DF22E2: call [004010ECh] ; rtcMidCharBstr
  loc_00DF22E8: mov var_48, esi
  loc_00DF22EB: mov esi, [00401228h] ; __vbaStrMove
  loc_00DF22F1: mov edx, eax
  loc_00DF22F3: lea ecx, var_18
  loc_00DF22F6: call __vbaStrMove
  loc_00DF22F8: push eax
  loc_00DF22F9: call [0040104Ch] ; rtcLowerCaseBstr
  loc_00DF22FF: mov edx, eax
  loc_00DF2301: lea ecx, var_1C
  loc_00DF2304: call __vbaStrMove
  loc_00DF2306: mov ecx, var_48
  loc_00DF2309: push eax
  loc_00DF230A: push edi
  loc_00DF230B: call [ecx+00000324h]
  loc_00DF2311: test eax, eax
  loc_00DF2313: fnclex
  loc_00DF2315: jge 00DF2329h
  loc_00DF2317: push 00000324h
  loc_00DF231C: push 006DCB00h
  loc_00DF2321: push edi
  loc_00DF2322: push eax
  loc_00DF2323: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2329: lea edx, var_1C
  loc_00DF232C: lea eax, var_18
  loc_00DF232F: push edx
  loc_00DF2330: push eax
  loc_00DF2331: push 00000002h
  loc_00DF2333: call [004011BCh] ; __vbaFreeStrList
  loc_00DF2339: add esp, 0000000Ch
  loc_00DF233C: lea ecx, var_2C
  loc_00DF233F: call [00401024h] ; __vbaFreeVar
  loc_00DF2345: push 00DF236Ah
  loc_00DF234A: jmp 00DF2369h
  loc_00DF234C: lea ecx, var_1C
  loc_00DF234F: lea edx, var_18
  loc_00DF2352: push ecx
  loc_00DF2353: push edx
  loc_00DF2354: push 00000002h
  loc_00DF2356: call [004011BCh] ; __vbaFreeStrList
  loc_00DF235C: add esp, 0000000Ch
  loc_00DF235F: lea ecx, var_2C
  loc_00DF2362: call [00401024h] ; __vbaFreeVar
  loc_00DF2368: ret
  loc_00DF2369: ret
  loc_00DF236A: mov ecx, var_10
  loc_00DF236D: pop edi
  loc_00DF236E: pop esi
  loc_00DF236F: xor eax, eax
  loc_00DF2371: mov fs:[00000000h], ecx
  loc_00DF2378: pop ebx
  loc_00DF2379: mov esp, ebp
  loc_00DF237B: pop ebp
  loc_00DF237C: retn 0004h
End Sub

Private Sub Proc_2_123_DF2390(arg_C, arg_10, arg_14) 'DF2390
  loc_00DF2390: push ecx
  loc_00DF2391: push ebx
  loc_00DF2392: push ebp
  loc_00DF2393: push esi
  loc_00DF2394: mov esi, arg_C
  loc_00DF2398: lea ecx, Proc_2_123_DF2390
  loc_00DF239C: push edi
  loc_00DF239D: mov eax, [esi]
  loc_00DF239F: push ecx
  loc_00DF23A0: push esi
  loc_00DF23A1: mov arg_10, 00000000h
  loc_00DF23A9: call [eax+00000108h]
  loc_00DF23AF: mov ebp, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF23B5: test eax, eax
  loc_00DF23B7: fnclex
  loc_00DF23B9: jge 00DF23C9h
  loc_00DF23BB: push 00000108h
  loc_00DF23C0: push 006DCB00h
  loc_00DF23C5: push esi
  loc_00DF23C6: push eax
  loc_00DF23C7: call ebp
  loc_00DF23C9: fld real4 ptr Me
  loc_00DF23CD: mov edi, [00401208h] ; __vbaFpI4
  loc_00DF23D3: call edi
  loc_00DF23D5: mov edx, [esi]
  loc_00DF23D7: mov [esi+000001D0h], eax
  loc_00DF23DD: lea eax, Me
  loc_00DF23E1: push eax
  loc_00DF23E2: push esi
  loc_00DF23E3: call [edx+00000100h]
  loc_00DF23E9: test eax, eax
  loc_00DF23EB: fnclex
  loc_00DF23ED: jge 00DF23FDh
  loc_00DF23EF: push 00000100h
  loc_00DF23F4: push 006DCB00h
  loc_00DF23F9: push esi
  loc_00DF23FA: push eax
  loc_00DF23FB: call ebp
  loc_00DF23FD: fld real4 ptr Me
  loc_00DF2401: call edi
  loc_00DF2403: mov ecx, [esi]
  loc_00DF2405: lea edx, Me
  loc_00DF2409: push edx
  loc_00DF240A: push esi
  loc_00DF240B: mov [esi+000001D4h], eax
  loc_00DF2411: call [ecx+000000D8h]
  loc_00DF2417: test eax, eax
  loc_00DF2419: fnclex
  loc_00DF241B: jge 00DF242Bh
  loc_00DF241D: push 000000D8h
  loc_00DF2422: push 006DCB00h
  loc_00DF2427: push esi
  loc_00DF2428: push eax
  loc_00DF2429: call ebp
  loc_00DF242B: mov edi, arg_14
  loc_00DF242F: mov ecx, Me
  loc_00DF2433: mov eax, [edi]
  loc_00DF2435: push eax
  loc_00DF2436: push 00000001h
  loc_00DF2438: push 00000001h
  loc_00DF243A: push ecx
  loc_00DF243B: call 006DD648h ; SetPixel(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2440: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00DF2446: call ebx
  loc_00DF2448: mov edx, [esi]
  loc_00DF244A: lea eax, Me
  loc_00DF244E: push eax
  loc_00DF244F: push esi
  loc_00DF2450: call [edx+000000D8h]
  loc_00DF2456: test eax, eax
  loc_00DF2458: fnclex
  loc_00DF245A: jge 00DF246Ah
  loc_00DF245C: push 000000D8h
  loc_00DF2461: push 006DCB00h
  loc_00DF2466: push esi
  loc_00DF2467: push eax
  loc_00DF2468: call ebp
  loc_00DF246A: mov edx, [esi+000001D0h]
  loc_00DF2470: mov ecx, [edi]
  loc_00DF2472: mov eax, Me
  loc_00DF2476: sub edx, 00000002h
  loc_00DF2479: push ecx
  loc_00DF247A: jo 00DF251Dh
  loc_00DF2480: push edx
  loc_00DF2481: push 00000001h
  loc_00DF2483: push eax
  loc_00DF2484: call 006DD648h ; SetPixel(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2489: call ebx
  loc_00DF248B: mov ecx, [esi]
  loc_00DF248D: lea edx, Me
  loc_00DF2491: push edx
  loc_00DF2492: push esi
  loc_00DF2493: call [ecx+000000D8h]
  loc_00DF2499: test eax, eax
  loc_00DF249B: fnclex
  loc_00DF249D: jge 00DF24ADh
  loc_00DF249F: push 000000D8h
  loc_00DF24A4: push 006DCB00h
  loc_00DF24A9: push esi
  loc_00DF24AA: push eax
  loc_00DF24AB: call ebp
  loc_00DF24AD: mov ecx, [esi+000001D4h]
  loc_00DF24B3: mov eax, [edi]
  loc_00DF24B5: mov edx, Me
  loc_00DF24B9: sub ecx, 00000002h
  loc_00DF24BC: push eax
  loc_00DF24BD: push 00000001h
  loc_00DF24BF: jo 00DF251Dh
  loc_00DF24C1: push ecx
  loc_00DF24C2: push edx
  loc_00DF24C3: call 006DD648h ; SetPixel(%x1v, %x2v, %x3v, %x4v)
  loc_00DF24C8: call ebx
  loc_00DF24CA: mov eax, [esi]
  loc_00DF24CC: lea ecx, Me
  loc_00DF24D0: push ecx
  loc_00DF24D1: push esi
  loc_00DF24D2: call [eax+000000D8h]
  loc_00DF24D8: test eax, eax
  loc_00DF24DA: fnclex
  loc_00DF24DC: jge 00DF24ECh
  loc_00DF24DE: push 000000D8h
  loc_00DF24E3: push 006DCB00h
  loc_00DF24E8: push esi
  loc_00DF24E9: push eax
  loc_00DF24EA: call ebp
  loc_00DF24EC: mov eax, [esi+000001D0h]
  loc_00DF24F2: mov ecx, [esi+000001D4h]
  loc_00DF24F8: mov edx, [edi]
  loc_00DF24FA: sub eax, 00000002h
  loc_00DF24FD: jo 00DF251Dh
  loc_00DF24FF: sub ecx, 00000002h
  loc_00DF2502: push edx
  loc_00DF2503: mov edx, arg_C
  loc_00DF2507: push eax
  loc_00DF2508: jo 00DF251Dh
  loc_00DF250A: push ecx
  loc_00DF250B: push edx
  loc_00DF250C: call 006DD648h ; SetPixel(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2511: call ebx
  loc_00DF2513: pop edi
  loc_00DF2514: pop esi
  loc_00DF2515: pop ebp
  loc_00DF2516: xor eax, eax
  loc_00DF2518: pop ebx
  loc_00DF2519: pop ecx
  loc_00DF251A: retn 0008h
End Sub

Private Sub Proc_2_124_DF2530(arg_C) 'DF2530
  loc_00DF2530: push ebp
  loc_00DF2531: mov ebp, esp
  loc_00DF2533: sub esp, 00000018h
  loc_00DF2536: push 00402836h ; __vbaExceptHandler
  loc_00DF253B: mov eax, fs:[00000000h]
  loc_00DF2541: push eax
  loc_00DF2542: mov fs:[00000000h], esp
  loc_00DF2549: mov eax, 000000B8h
  loc_00DF254E: call 00402830h ; __vbaChkstk
  loc_00DF2553: push ebx
  loc_00DF2554: push esi
  loc_00DF2555: push edi
  loc_00DF2556: mov var_18, esp
  loc_00DF2559: mov var_14, 00401438h ; "$"
  loc_00DF2560: mov var_10, 00000000h
  loc_00DF2567: mov var_C, 00000000h
  loc_00DF256E: mov var_4, 00000001h
  loc_00DF2575: mov var_4, 00000002h
  loc_00DF257C: lea eax, var_4C
  loc_00DF257F: push eax
  loc_00DF2580: mov ecx, Me
  loc_00DF2583: mov edx, [ecx]
  loc_00DF2585: mov eax, Me
  loc_00DF2588: push eax
  loc_00DF2589: call [edx+00000108h]
  loc_00DF258F: fnclex
  loc_00DF2591: mov var_60, eax
  loc_00DF2594: cmp var_60, 00000000h
  loc_00DF2598: jge 00DF25BAh
  loc_00DF259A: push 00000108h
  loc_00DF259F: push 006DCB00h
  loc_00DF25A4: mov ecx, Me
  loc_00DF25A7: push ecx
  loc_00DF25A8: mov edx, var_60
  loc_00DF25AB: push edx
  loc_00DF25AC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF25B2: mov var_94, eax
  loc_00DF25B8: jmp 00DF25C4h
  loc_00DF25BA: mov var_94, 00000000h
  loc_00DF25C4: fld real4 ptr var_4C
  loc_00DF25C7: call [00401208h] ; __vbaFpI4
  loc_00DF25CD: mov ecx, Me
  loc_00DF25D0: mov [ecx+000001D0h], eax
  loc_00DF25D6: mov var_4, 00000003h
  loc_00DF25DD: lea edx, var_4C
  loc_00DF25E0: push edx
  loc_00DF25E1: mov eax, Me
  loc_00DF25E4: mov ecx, [eax]
  loc_00DF25E6: mov edx, Me
  loc_00DF25E9: push edx
  loc_00DF25EA: call [ecx+00000100h]
  loc_00DF25F0: fnclex
  loc_00DF25F2: mov var_60, eax
  loc_00DF25F5: cmp var_60, 00000000h
  loc_00DF25F9: jge 00DF261Bh
  loc_00DF25FB: push 00000100h
  loc_00DF2600: push 006DCB00h
  loc_00DF2605: mov eax, Me
  loc_00DF2608: push eax
  loc_00DF2609: mov ecx, var_60
  loc_00DF260C: push ecx
  loc_00DF260D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2613: mov var_98, eax
  loc_00DF2619: jmp 00DF2625h
  loc_00DF261B: mov var_98, 00000000h
  loc_00DF2625: fld real4 ptr var_4C
  loc_00DF2628: call [00401208h] ; __vbaFpI4
  loc_00DF262E: mov edx, Me
  loc_00DF2631: mov [edx+000001D4h], eax
  loc_00DF2637: mov var_4, 00000004h
  loc_00DF263E: mov eax, Me
  loc_00DF2641: mov ecx, [eax+000001D0h]
  loc_00DF2647: push ecx
  loc_00DF2648: mov edx, Me
  loc_00DF264B: mov eax, [edx+000001D4h]
  loc_00DF2651: push eax
  loc_00DF2652: push 00000000h
  loc_00DF2654: push 00000000h
  loc_00DF2656: mov ecx, Me
  loc_00DF2659: add ecx, 000000ACh
  loc_00DF265F: push ecx
  loc_00DF2660: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF2665: call [00401074h] ; __vbaSetSystemError
  loc_00DF266B: mov var_4, 00000005h
  loc_00DF2672: mov edx, Me
  loc_00DF2675: movsx eax, [edx+0000007Ch]
  loc_00DF2679: test eax, eax
  loc_00DF267B: jnz 00DF2796h
  loc_00DF2681: mov var_4, 00000006h
  loc_00DF2688: mov ecx, Me
  loc_00DF268B: cmp [ecx+00000044h], 00000000h
  loc_00DF268F: jnz 00DF26FFh
  loc_00DF2691: mov var_4, 00000007h
  loc_00DF2698: lea edx, var_4C
  loc_00DF269B: push edx
  loc_00DF269C: mov eax, Me
  loc_00DF269F: mov ecx, [eax]
  loc_00DF26A1: mov edx, Me
  loc_00DF26A4: push edx
  loc_00DF26A5: call [ecx+000000D8h]
  loc_00DF26AB: fnclex
  loc_00DF26AD: mov var_60, eax
  loc_00DF26B0: cmp var_60, 00000000h
  loc_00DF26B4: jge 00DF26D6h
  loc_00DF26B6: push 000000D8h
  loc_00DF26BB: push 006DCB00h
  loc_00DF26C0: mov eax, Me
  loc_00DF26C3: push eax
  loc_00DF26C4: mov ecx, var_60
  loc_00DF26C7: push ecx
  loc_00DF26C8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF26CE: mov var_9C, eax
  loc_00DF26D4: jmp 00DF26E0h
  loc_00DF26D6: mov var_9C, 00000000h
  loc_00DF26E0: push 0000000Fh
  loc_00DF26E2: push 00000005h
  loc_00DF26E4: mov edx, Me
  loc_00DF26E7: add edx, 000000ACh
  loc_00DF26ED: push edx
  loc_00DF26EE: mov eax, var_4C
  loc_00DF26F1: push eax
  loc_00DF26F2: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF26F7: call [00401074h] ; __vbaSetSystemError
  loc_00DF26FD: jmp 00DF277Bh
  loc_00DF26FF: mov var_4, 00000008h
  loc_00DF2706: mov ecx, Me
  loc_00DF2709: cmp [ecx+00000044h], 00000001h
  loc_00DF270D: jnz 00DF277Bh
  loc_00DF270F: mov var_4, 00000009h
  loc_00DF2716: lea edx, var_4C
  loc_00DF2719: push edx
  loc_00DF271A: mov eax, Me
  loc_00DF271D: mov ecx, [eax]
  loc_00DF271F: mov edx, Me
  loc_00DF2722: push edx
  loc_00DF2723: call [ecx+000000D8h]
  loc_00DF2729: fnclex
  loc_00DF272B: mov var_60, eax
  loc_00DF272E: cmp var_60, 00000000h
  loc_00DF2732: jge 00DF2754h
  loc_00DF2734: push 000000D8h
  loc_00DF2739: push 006DCB00h
  loc_00DF273E: mov eax, Me
  loc_00DF2741: push eax
  loc_00DF2742: mov ecx, var_60
  loc_00DF2745: push ecx
  loc_00DF2746: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF274C: mov var_A0, eax
  loc_00DF2752: jmp 00DF275Eh
  loc_00DF2754: mov var_A0, 00000000h
  loc_00DF275E: push 0000000Fh
  loc_00DF2760: push 00000004h
  loc_00DF2762: mov edx, Me
  loc_00DF2765: add edx, 000000ACh
  loc_00DF276B: push edx
  loc_00DF276C: mov eax, var_4C
  loc_00DF276F: push eax
  loc_00DF2770: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2775: call [00401074h] ; __vbaSetSystemError
  loc_00DF277B: mov var_4, 0000000Bh
  loc_00DF2782: mov ecx, Me
  loc_00DF2785: mov edx, [ecx]
  loc_00DF2787: mov eax, Me
  loc_00DF278A: push eax
  loc_00DF278B: call [edx+00000920h]
  loc_00DF2791: jmp 00DF3297h
  loc_00DF2796: mov var_4, 0000000Eh
  loc_00DF279D: mov ecx, Me
  loc_00DF27A0: cmp [ecx+00000064h], 00000000h
  loc_00DF27A4: jz 00DF2953h
  loc_00DF27AA: mov edx, Me
  loc_00DF27AD: movsx eax, [edx+0000006Ch]
  loc_00DF27B1: test eax, eax
  loc_00DF27B3: jz 00DF2953h
  loc_00DF27B9: mov var_4, 0000000Fh
  loc_00DF27C0: mov var_4C, 00000000h
  loc_00DF27C7: lea ecx, var_50
  loc_00DF27CA: push ecx
  loc_00DF27CB: lea edx, var_4C
  loc_00DF27CE: push edx
  loc_00DF27CF: mov eax, Me
  loc_00DF27D2: mov ecx, [eax+00000088h]
  loc_00DF27D8: push ecx
  loc_00DF27D9: mov edx, Me
  loc_00DF27DC: mov eax, [edx]
  loc_00DF27DE: mov ecx, Me
  loc_00DF27E1: push ecx
  loc_00DF27E2: call [eax+0000090Ch]
  loc_00DF27E8: mov var_58, 3CA3D70Ah
  loc_00DF27EF: mov edx, var_50
  loc_00DF27F2: mov var_54, edx
  loc_00DF27F5: lea eax, var_5C
  loc_00DF27F8: push eax
  loc_00DF27F9: lea ecx, var_58
  loc_00DF27FC: push ecx
  loc_00DF27FD: lea edx, var_54
  loc_00DF2800: push edx
  loc_00DF2801: mov eax, Me
  loc_00DF2804: mov ecx, [eax]
  loc_00DF2806: mov edx, Me
  loc_00DF2809: push edx
  loc_00DF280A: call [ecx+0000097Ch]
  loc_00DF2810: mov eax, Me
  loc_00DF2813: add eax, 000000ACh
  loc_00DF2818: push eax
  loc_00DF2819: mov ecx, var_5C
  loc_00DF281C: push ecx
  loc_00DF281D: mov edx, Me
  loc_00DF2820: mov eax, [edx]
  loc_00DF2822: mov ecx, Me
  loc_00DF2825: push ecx
  loc_00DF2826: call [eax+00000974h]
  loc_00DF282C: mov var_4, 00000010h
  loc_00DF2833: mov edx, Me
  loc_00DF2836: mov eax, [edx]
  loc_00DF2838: mov ecx, Me
  loc_00DF283B: push ecx
  loc_00DF283C: call [eax+00000920h]
  loc_00DF2842: mov var_4, 00000011h
  loc_00DF2849: mov edx, Me
  loc_00DF284C: cmp [edx+00000044h], 0000000Ch
  loc_00DF2850: jz 00DF294Eh
  loc_00DF2856: mov var_4, 00000012h
  loc_00DF285D: lea eax, var_4C
  loc_00DF2860: push eax
  loc_00DF2861: mov ecx, Me
  loc_00DF2864: mov edx, [ecx]
  loc_00DF2866: mov eax, Me
  loc_00DF2869: push eax
  loc_00DF286A: call [edx+000000D8h]
  loc_00DF2870: fnclex
  loc_00DF2872: mov var_60, eax
  loc_00DF2875: cmp var_60, 00000000h
  loc_00DF2879: jge 00DF289Bh
  loc_00DF287B: push 000000D8h
  loc_00DF2880: push 006DCB00h
  loc_00DF2885: mov ecx, Me
  loc_00DF2888: push ecx
  loc_00DF2889: mov edx, var_60
  loc_00DF288C: push edx
  loc_00DF288D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2893: mov var_A4, eax
  loc_00DF2899: jmp 00DF28A5h
  loc_00DF289B: mov var_A4, 00000000h
  loc_00DF28A5: push 0000000Fh
  loc_00DF28A7: push 0000000Ah
  loc_00DF28A9: mov eax, Me
  loc_00DF28AC: add eax, 000000ACh
  loc_00DF28B1: push eax
  loc_00DF28B2: mov ecx, var_4C
  loc_00DF28B5: push ecx
  loc_00DF28B6: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF28BB: call [00401074h] ; __vbaSetSystemError
  loc_00DF28C1: mov var_4, 00000013h
  loc_00DF28C8: mov edx, Me
  loc_00DF28CB: movsx eax, [edx+0000006Eh]
  loc_00DF28CF: test eax, eax
  loc_00DF28D1: jz 00DF294Eh
  loc_00DF28D3: mov ecx, Me
  loc_00DF28D6: movsx edx, [ecx+00000050h]
  loc_00DF28DA: test edx, edx
  loc_00DF28DC: jz 00DF294Eh
  loc_00DF28DE: mov eax, Me
  loc_00DF28E1: cmp [eax+00000044h], 00000000h
  loc_00DF28E5: jnz 00DF294Eh
  loc_00DF28E7: mov var_4, 00000014h
  loc_00DF28EE: mov var_4C, 00000000h
  loc_00DF28F5: lea ecx, var_50
  loc_00DF28F8: push ecx
  loc_00DF28F9: lea edx, var_4C
  loc_00DF28FC: push edx
  loc_00DF28FD: push 8000000Ch
  loc_00DF2902: mov eax, Me
  loc_00DF2905: mov ecx, [eax]
  loc_00DF2907: mov edx, Me
  loc_00DF290A: push edx
  loc_00DF290B: call [ecx+0000090Ch]
  loc_00DF2911: mov eax, var_50
  loc_00DF2914: push eax
  loc_00DF2915: mov ecx, Me
  loc_00DF2918: mov edx, [ecx+000001D0h]
  loc_00DF291E: sub edx, 00000007h
  loc_00DF2921: jo 00DF32DCh
  loc_00DF2927: push edx
  loc_00DF2928: mov eax, Me
  loc_00DF292B: mov ecx, [eax+000001D4h]
  loc_00DF2931: sub ecx, 00000007h
  loc_00DF2934: jo 00DF32DCh
  loc_00DF293A: push ecx
  loc_00DF293B: push 00000004h
  loc_00DF293D: push 00000004h
  loc_00DF293F: mov edx, Me
  loc_00DF2942: mov eax, [edx]
  loc_00DF2944: mov ecx, Me
  loc_00DF2947: push ecx
  loc_00DF2948: call [eax+000008F0h]
  loc_00DF294E: jmp 00DF3297h
  loc_00DF2953: mov var_4, 00000019h
  loc_00DF295A: mov edx, arg_C
  loc_00DF295D: mov var_70, edx
  loc_00DF2960: mov eax, var_70
  loc_00DF2963: mov var_A8, eax
  loc_00DF2969: cmp var_A8, 00000000h
  loc_00DF2970: jz 00DF2996h
  loc_00DF2972: cmp var_A8, 00000001h
  loc_00DF2979: jz 00DF2B21h
  loc_00DF297F: cmp var_A8, 00000002h
  loc_00DF2986: jz 00DF2C93h
  loc_00DF298C: jmp 00DF2F20h
  loc_00DF2991: jmp 00DF2F20h
  loc_00DF2996: mov var_4, 0000001Bh
  loc_00DF299D: mov ecx, Me
  loc_00DF29A0: mov edx, [ecx]
  loc_00DF29A2: mov eax, Me
  loc_00DF29A5: push eax
  loc_00DF29A6: call [edx+00000914h]
  loc_00DF29AC: mov var_4, 0000001Ch
  loc_00DF29B3: mov var_4C, 00000000h
  loc_00DF29BA: lea ecx, var_50
  loc_00DF29BD: push ecx
  loc_00DF29BE: lea edx, var_4C
  loc_00DF29C1: push edx
  loc_00DF29C2: mov eax, Me
  loc_00DF29C5: mov ecx, [eax+00000088h]
  loc_00DF29CB: push ecx
  loc_00DF29CC: mov edx, Me
  loc_00DF29CF: mov eax, [edx]
  loc_00DF29D1: mov ecx, Me
  loc_00DF29D4: push ecx
  loc_00DF29D5: call [eax+0000090Ch]
  loc_00DF29DB: mov edx, Me
  loc_00DF29DE: add edx, 000000ACh
  loc_00DF29E4: push edx
  loc_00DF29E5: mov eax, var_50
  loc_00DF29E8: push eax
  loc_00DF29E9: mov ecx, Me
  loc_00DF29EC: mov edx, [ecx]
  loc_00DF29EE: mov eax, Me
  loc_00DF29F1: push eax
  loc_00DF29F2: call [edx+00000974h]
  loc_00DF29F8: mov var_4, 0000001Dh
  loc_00DF29FF: mov ecx, Me
  loc_00DF2A02: mov edx, [ecx]
  loc_00DF2A04: mov eax, Me
  loc_00DF2A07: push eax
  loc_00DF2A08: call [edx+00000920h]
  loc_00DF2A0E: mov var_4, 0000001Eh
  loc_00DF2A15: mov ecx, Me
  loc_00DF2A18: mov edx, [ecx+00000044h]
  loc_00DF2A1B: mov var_74, edx
  loc_00DF2A1E: mov eax, var_74
  loc_00DF2A21: mov var_AC, eax
  loc_00DF2A27: cmp var_AC, 00000000h
  loc_00DF2A2E: jz 00DF2A43h
  loc_00DF2A30: cmp var_AC, 00000001h
  loc_00DF2A37: jz 00DF2AB1h
  loc_00DF2A39: jmp 00DF2B1Ch
  loc_00DF2A3E: jmp 00DF2B1Ch
  loc_00DF2A43: mov var_4, 00000020h
  loc_00DF2A4A: lea ecx, var_4C
  loc_00DF2A4D: push ecx
  loc_00DF2A4E: mov edx, Me
  loc_00DF2A51: mov eax, [edx]
  loc_00DF2A53: mov ecx, Me
  loc_00DF2A56: push ecx
  loc_00DF2A57: call [eax+000000D8h]
  loc_00DF2A5D: fnclex
  loc_00DF2A5F: mov var_60, eax
  loc_00DF2A62: cmp var_60, 00000000h
  loc_00DF2A66: jge 00DF2A88h
  loc_00DF2A68: push 000000D8h
  loc_00DF2A6D: push 006DCB00h
  loc_00DF2A72: mov edx, Me
  loc_00DF2A75: push edx
  loc_00DF2A76: mov eax, var_60
  loc_00DF2A79: push eax
  loc_00DF2A7A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2A80: mov var_B0, eax
  loc_00DF2A86: jmp 00DF2A92h
  loc_00DF2A88: mov var_B0, 00000000h
  loc_00DF2A92: push 0000000Fh
  loc_00DF2A94: push 00000005h
  loc_00DF2A96: mov ecx, Me
  loc_00DF2A99: add ecx, 000000ACh
  loc_00DF2A9F: push ecx
  loc_00DF2AA0: mov edx, var_4C
  loc_00DF2AA3: push edx
  loc_00DF2AA4: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2AA9: call [00401074h] ; __vbaSetSystemError
  loc_00DF2AAF: jmp 00DF2B1Ch
  loc_00DF2AB1: mov var_4, 00000022h
  loc_00DF2AB8: lea eax, var_4C
  loc_00DF2ABB: push eax
  loc_00DF2ABC: mov ecx, Me
  loc_00DF2ABF: mov edx, [ecx]
  loc_00DF2AC1: mov eax, Me
  loc_00DF2AC4: push eax
  loc_00DF2AC5: call [edx+000000D8h]
  loc_00DF2ACB: fnclex
  loc_00DF2ACD: mov var_60, eax
  loc_00DF2AD0: cmp var_60, 00000000h
  loc_00DF2AD4: jge 00DF2AF6h
  loc_00DF2AD6: push 000000D8h
  loc_00DF2ADB: push 006DCB00h
  loc_00DF2AE0: mov ecx, Me
  loc_00DF2AE3: push ecx
  loc_00DF2AE4: mov edx, var_60
  loc_00DF2AE7: push edx
  loc_00DF2AE8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2AEE: mov var_B4, eax
  loc_00DF2AF4: jmp 00DF2B00h
  loc_00DF2AF6: mov var_B4, 00000000h
  loc_00DF2B00: push 0000000Fh
  loc_00DF2B02: push 00000004h
  loc_00DF2B04: mov eax, Me
  loc_00DF2B07: add eax, 000000ACh
  loc_00DF2B0C: push eax
  loc_00DF2B0D: mov ecx, var_4C
  loc_00DF2B10: push ecx
  loc_00DF2B11: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2B16: call [00401074h] ; __vbaSetSystemError
  loc_00DF2B1C: jmp 00DF2F20h
  loc_00DF2B21: mov var_4, 00000025h
  loc_00DF2B28: mov var_4C, 00000000h
  loc_00DF2B2F: lea edx, var_50
  loc_00DF2B32: push edx
  loc_00DF2B33: lea eax, var_4C
  loc_00DF2B36: push eax
  loc_00DF2B37: mov ecx, Me
  loc_00DF2B3A: mov edx, [ecx+00000088h]
  loc_00DF2B40: push edx
  loc_00DF2B41: mov eax, Me
  loc_00DF2B44: mov ecx, [eax]
  loc_00DF2B46: mov edx, Me
  loc_00DF2B49: push edx
  loc_00DF2B4A: call [ecx+0000090Ch]
  loc_00DF2B50: mov eax, Me
  loc_00DF2B53: add eax, 000000ACh
  loc_00DF2B58: push eax
  loc_00DF2B59: mov ecx, var_50
  loc_00DF2B5C: push ecx
  loc_00DF2B5D: mov edx, Me
  loc_00DF2B60: mov eax, [edx]
  loc_00DF2B62: mov ecx, Me
  loc_00DF2B65: push ecx
  loc_00DF2B66: call [eax+00000974h]
  loc_00DF2B6C: mov var_4, 00000026h
  loc_00DF2B73: mov edx, Me
  loc_00DF2B76: mov eax, [edx]
  loc_00DF2B78: mov ecx, Me
  loc_00DF2B7B: push ecx
  loc_00DF2B7C: call [eax+00000920h]
  loc_00DF2B82: mov var_4, 00000027h
  loc_00DF2B89: mov edx, Me
  loc_00DF2B8C: mov eax, [edx+00000044h]
  loc_00DF2B8F: mov var_78, eax
  loc_00DF2B92: mov ecx, var_78
  loc_00DF2B95: mov var_B8, ecx
  loc_00DF2B9B: cmp var_B8, 00000001h
  loc_00DF2BA2: jz 00DF2BB4h
  loc_00DF2BA4: cmp var_B8, 0000000Ch
  loc_00DF2BAB: jz 00DF2BB4h
  loc_00DF2BAD: jmp 00DF2C22h
  loc_00DF2BAF: jmp 00DF2C8Eh
  loc_00DF2BB4: mov var_4, 00000029h
  loc_00DF2BBB: lea edx, var_4C
  loc_00DF2BBE: push edx
  loc_00DF2BBF: mov eax, Me
  loc_00DF2BC2: mov ecx, [eax]
  loc_00DF2BC4: mov edx, Me
  loc_00DF2BC7: push edx
  loc_00DF2BC8: call [ecx+000000D8h]
  loc_00DF2BCE: fnclex
  loc_00DF2BD0: mov var_60, eax
  loc_00DF2BD3: cmp var_60, 00000000h
  loc_00DF2BD7: jge 00DF2BF9h
  loc_00DF2BD9: push 000000D8h
  loc_00DF2BDE: push 006DCB00h
  loc_00DF2BE3: mov eax, Me
  loc_00DF2BE6: push eax
  loc_00DF2BE7: mov ecx, var_60
  loc_00DF2BEA: push ecx
  loc_00DF2BEB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2BF1: mov var_BC, eax
  loc_00DF2BF7: jmp 00DF2C03h
  loc_00DF2BF9: mov var_BC, 00000000h
  loc_00DF2C03: push 0000000Fh
  loc_00DF2C05: push 00000004h
  loc_00DF2C07: mov edx, Me
  loc_00DF2C0A: add edx, 000000ACh
  loc_00DF2C10: push edx
  loc_00DF2C11: mov eax, var_4C
  loc_00DF2C14: push eax
  loc_00DF2C15: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2C1A: call [00401074h] ; __vbaSetSystemError
  loc_00DF2C20: jmp 00DF2C8Eh
  loc_00DF2C22: mov var_4, 0000002Bh
  loc_00DF2C29: lea ecx, var_4C
  loc_00DF2C2C: push ecx
  loc_00DF2C2D: mov edx, Me
  loc_00DF2C30: mov eax, [edx]
  loc_00DF2C32: mov ecx, Me
  loc_00DF2C35: push ecx
  loc_00DF2C36: call [eax+000000D8h]
  loc_00DF2C3C: fnclex
  loc_00DF2C3E: mov var_60, eax
  loc_00DF2C41: cmp var_60, 00000000h
  loc_00DF2C45: jge 00DF2C67h
  loc_00DF2C47: push 000000D8h
  loc_00DF2C4C: push 006DCB00h
  loc_00DF2C51: mov edx, Me
  loc_00DF2C54: push edx
  loc_00DF2C55: mov eax, var_60
  loc_00DF2C58: push eax
  loc_00DF2C59: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2C5F: mov var_C0, eax
  loc_00DF2C65: jmp 00DF2C71h
  loc_00DF2C67: mov var_C0, 00000000h
  loc_00DF2C71: push 0000000Fh
  loc_00DF2C73: push 00000005h
  loc_00DF2C75: mov ecx, Me
  loc_00DF2C78: add ecx, 000000ACh
  loc_00DF2C7E: push ecx
  loc_00DF2C7F: mov edx, var_4C
  loc_00DF2C82: push edx
  loc_00DF2C83: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2C88: call [00401074h] ; __vbaSetSystemError
  loc_00DF2C8E: jmp 00DF2F20h
  loc_00DF2C93: mov var_4, 0000002Eh
  loc_00DF2C9A: mov var_4C, 00000000h
  loc_00DF2CA1: lea eax, var_50
  loc_00DF2CA4: push eax
  loc_00DF2CA5: lea ecx, var_4C
  loc_00DF2CA8: push ecx
  loc_00DF2CA9: mov edx, Me
  loc_00DF2CAC: mov eax, [edx+00000088h]
  loc_00DF2CB2: push eax
  loc_00DF2CB3: mov ecx, Me
  loc_00DF2CB6: mov edx, [ecx]
  loc_00DF2CB8: mov eax, Me
  loc_00DF2CBB: push eax
  loc_00DF2CBC: call [edx+0000090Ch]
  loc_00DF2CC2: mov ecx, Me
  loc_00DF2CC5: add ecx, 000000ACh
  loc_00DF2CCB: push ecx
  loc_00DF2CCC: mov edx, var_50
  loc_00DF2CCF: push edx
  loc_00DF2CD0: mov eax, Me
  loc_00DF2CD3: mov ecx, [eax]
  loc_00DF2CD5: mov edx, Me
  loc_00DF2CD8: push edx
  loc_00DF2CD9: call [ecx+00000974h]
  loc_00DF2CDF: mov var_4, 0000002Fh
  loc_00DF2CE6: mov eax, Me
  loc_00DF2CE9: mov ecx, [eax]
  loc_00DF2CEB: mov edx, Me
  loc_00DF2CEE: push edx
  loc_00DF2CEF: call [ecx+00000920h]
  loc_00DF2CF5: mov var_4, 00000030h
  loc_00DF2CFC: mov eax, Me
  loc_00DF2CFF: mov ecx, [eax+00000044h]
  loc_00DF2D02: mov var_7C, ecx
  loc_00DF2D05: mov edx, var_7C
  loc_00DF2D08: mov var_C4, edx
  loc_00DF2D0E: cmp var_C4, 0000000Ch
  loc_00DF2D15: ja 00DF2F20h
  loc_00DF2D1B: mov ecx, var_C4
  loc_00DF2D21: xor eax, eax
  loc_00DF2D23: mov al, [ecx+00DF32CFh]
  loc_00DF2D29: jmp [eax*4+00DF32BFh]
  loc_00DF2D30: jmp 00DF2F20h
  loc_00DF2D35: mov var_4, 00000032h
  loc_00DF2D3C: mov var_4C, 00000000h
  loc_00DF2D43: lea edx, var_50
  loc_00DF2D46: push edx
  loc_00DF2D47: lea eax, var_4C
  loc_00DF2D4A: push eax
  loc_00DF2D4B: push 0099A8ACh
  loc_00DF2D50: mov ecx, Me
  loc_00DF2D53: mov edx, [ecx]
  loc_00DF2D55: mov eax, Me
  loc_00DF2D58: push eax
  loc_00DF2D59: call [edx+0000090Ch]
  loc_00DF2D5F: mov ecx, var_50
  loc_00DF2D62: push ecx
  loc_00DF2D63: mov edx, Me
  loc_00DF2D66: mov eax, [edx+000001D0h]
  loc_00DF2D6C: sub eax, 00000002h
  loc_00DF2D6F: jo 00DF32DCh
  loc_00DF2D75: push eax
  loc_00DF2D76: mov ecx, Me
  loc_00DF2D79: mov edx, [ecx+000001D4h]
  loc_00DF2D7F: sub edx, 00000002h
  loc_00DF2D82: jo 00DF32DCh
  loc_00DF2D88: push edx
  loc_00DF2D89: push 00000001h
  loc_00DF2D8B: push 00000001h
  loc_00DF2D8D: mov eax, Me
  loc_00DF2D90: mov ecx, [eax]
  loc_00DF2D92: mov edx, Me
  loc_00DF2D95: push edx
  loc_00DF2D96: call [ecx+000008F0h]
  loc_00DF2D9C: mov var_4, 00000033h
  loc_00DF2DA3: mov var_4C, 00000000h
  loc_00DF2DAA: lea eax, var_50
  loc_00DF2DAD: push eax
  loc_00DF2DAE: lea ecx, var_4C
  loc_00DF2DB1: push ecx
  loc_00DF2DB2: push 00000000h
  loc_00DF2DB4: mov edx, Me
  loc_00DF2DB7: mov eax, [edx]
  loc_00DF2DB9: mov ecx, Me
  loc_00DF2DBC: push ecx
  loc_00DF2DBD: call [eax+0000090Ch]
  loc_00DF2DC3: mov edx, var_50
  loc_00DF2DC6: push edx
  loc_00DF2DC7: mov eax, Me
  loc_00DF2DCA: mov ecx, [eax+000001D0h]
  loc_00DF2DD0: push ecx
  loc_00DF2DD1: mov edx, Me
  loc_00DF2DD4: mov eax, [edx+000001D4h]
  loc_00DF2DDA: push eax
  loc_00DF2DDB: push 00000000h
  loc_00DF2DDD: push 00000000h
  loc_00DF2DDF: mov ecx, Me
  loc_00DF2DE2: mov edx, [ecx]
  loc_00DF2DE4: mov eax, Me
  loc_00DF2DE7: push eax
  loc_00DF2DE8: call [edx+000008F0h]
  loc_00DF2DEE: jmp 00DF2F20h
  loc_00DF2DF3: mov var_4, 00000035h
  loc_00DF2DFA: lea ecx, var_4C
  loc_00DF2DFD: push ecx
  loc_00DF2DFE: mov edx, Me
  loc_00DF2E01: mov eax, [edx]
  loc_00DF2E03: mov ecx, Me
  loc_00DF2E06: push ecx
  loc_00DF2E07: call [eax+000000D8h]
  loc_00DF2E0D: fnclex
  loc_00DF2E0F: mov var_60, eax
  loc_00DF2E12: cmp var_60, 00000000h
  loc_00DF2E16: jge 00DF2E38h
  loc_00DF2E18: push 000000D8h
  loc_00DF2E1D: push 006DCB00h
  loc_00DF2E22: mov edx, Me
  loc_00DF2E25: push edx
  loc_00DF2E26: mov eax, var_60
  loc_00DF2E29: push eax
  loc_00DF2E2A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2E30: mov var_C8, eax
  loc_00DF2E36: jmp 00DF2E42h
  loc_00DF2E38: mov var_C8, 00000000h
  loc_00DF2E42: push 0000000Fh
  loc_00DF2E44: push 0000000Ah
  loc_00DF2E46: mov ecx, Me
  loc_00DF2E49: add ecx, 000000ACh
  loc_00DF2E4F: push ecx
  loc_00DF2E50: mov edx, var_4C
  loc_00DF2E53: push edx
  loc_00DF2E54: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF2E59: call [00401074h] ; __vbaSetSystemError
  loc_00DF2E5F: jmp 00DF2F20h
  loc_00DF2E64: mov var_4, 00000037h
  loc_00DF2E6B: mov var_4C, 00000000h
  loc_00DF2E72: lea eax, var_50
  loc_00DF2E75: push eax
  loc_00DF2E76: lea ecx, var_4C
  loc_00DF2E79: push ecx
  loc_00DF2E7A: push 00FFFFFFh
  loc_00DF2E7F: mov edx, Me
  loc_00DF2E82: mov eax, [edx]
  loc_00DF2E84: mov ecx, Me
  loc_00DF2E87: push ecx
  loc_00DF2E88: call [eax+0000090Ch]
  loc_00DF2E8E: mov edx, var_50
  loc_00DF2E91: push edx
  loc_00DF2E92: mov eax, Me
  loc_00DF2E95: mov ecx, [eax+000001D0h]
  loc_00DF2E9B: push ecx
  loc_00DF2E9C: mov edx, Me
  loc_00DF2E9F: mov eax, [edx+000001D4h]
  loc_00DF2EA5: push eax
  loc_00DF2EA6: push 00000000h
  loc_00DF2EA8: push 00000000h
  loc_00DF2EAA: mov ecx, Me
  loc_00DF2EAD: mov edx, [ecx]
  loc_00DF2EAF: mov eax, Me
  loc_00DF2EB2: push eax
  loc_00DF2EB3: call [edx+000008F0h]
  loc_00DF2EB9: mov var_4, 00000038h
  loc_00DF2EC0: mov var_4C, 00000000h
  loc_00DF2EC7: lea ecx, var_50
  loc_00DF2ECA: push ecx
  loc_00DF2ECB: lea edx, var_4C
  loc_00DF2ECE: push edx
  loc_00DF2ECF: push 80000011h
  loc_00DF2ED4: mov eax, Me
  loc_00DF2ED7: mov ecx, [eax]
  loc_00DF2ED9: mov edx, Me
  loc_00DF2EDC: push edx
  loc_00DF2EDD: call [ecx+0000090Ch]
  loc_00DF2EE3: mov eax, var_50
  loc_00DF2EE6: push eax
  loc_00DF2EE7: mov ecx, Me
  loc_00DF2EEA: mov edx, [ecx+000001D0h]
  loc_00DF2EF0: add edx, 00000001h
  loc_00DF2EF3: jo 00DF32DCh
  loc_00DF2EF9: push edx
  loc_00DF2EFA: mov eax, Me
  loc_00DF2EFD: mov ecx, [eax+000001D4h]
  loc_00DF2F03: add ecx, 00000001h
  loc_00DF2F06: jo 00DF32DCh
  loc_00DF2F0C: push ecx
  loc_00DF2F0D: push 00000000h
  loc_00DF2F0F: push 00000000h
  loc_00DF2F11: mov edx, Me
  loc_00DF2F14: mov eax, [edx]
  loc_00DF2F16: mov ecx, Me
  loc_00DF2F19: push ecx
  loc_00DF2F1A: call [eax+000008F0h]
  loc_00DF2F20: mov var_4, 0000003Bh
  loc_00DF2F27: mov edx, Me
  loc_00DF2F2A: movsx eax, [edx+00000050h]
  loc_00DF2F2E: test eax, eax
  loc_00DF2F30: jnz 00DF2F41h
  loc_00DF2F32: mov ecx, Me
  loc_00DF2F35: movsx edx, [ecx+00000058h]
  loc_00DF2F39: test edx, edx
  loc_00DF2F3B: jz 00DF3297h
  loc_00DF2F41: mov var_4, 0000003Ch
  loc_00DF2F48: push FFFFFFFFh
  loc_00DF2F4A: call [004010A4h] ; __vbaOnError
  loc_00DF2F50: mov var_4, 0000003Dh
  loc_00DF2F57: lea eax, var_44
  loc_00DF2F5A: push eax
  loc_00DF2F5B: mov ecx, Me
  loc_00DF2F5E: mov edx, [ecx]
  loc_00DF2F60: mov eax, Me
  loc_00DF2F63: push eax
  loc_00DF2F64: call [edx+000002B0h]
  loc_00DF2F6A: fnclex
  loc_00DF2F6C: mov var_60, eax
  loc_00DF2F6F: cmp var_60, 00000000h
  loc_00DF2F73: jge 00DF2F95h
  loc_00DF2F75: push 000002B0h
  loc_00DF2F7A: push 006DCB00h
  loc_00DF2F7F: mov ecx, Me
  loc_00DF2F82: push ecx
  loc_00DF2F83: mov edx, var_60
  loc_00DF2F86: push edx
  loc_00DF2F87: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2F8D: mov var_CC, eax
  loc_00DF2F93: jmp 00DF2F9Fh
  loc_00DF2F95: mov var_CC, 00000000h
  loc_00DF2F9F: mov eax, var_44
  loc_00DF2FA2: mov var_64, eax
  loc_00DF2FA5: lea ecx, var_48
  loc_00DF2FA8: push ecx
  loc_00DF2FA9: mov edx, var_64
  loc_00DF2FAC: mov eax, [edx]
  loc_00DF2FAE: mov ecx, var_64
  loc_00DF2FB1: push ecx
  loc_00DF2FB2: call [eax+0000003Ch]
  loc_00DF2FB5: fnclex
  loc_00DF2FB7: mov var_68, eax
  loc_00DF2FBA: cmp var_68, 00000000h
  loc_00DF2FBE: jge 00DF2FDDh
  loc_00DF2FC0: push 0000003Ch
  loc_00DF2FC2: push 006DEA70h
  loc_00DF2FC7: mov edx, var_64
  loc_00DF2FCA: push edx
  loc_00DF2FCB: mov eax, var_68
  loc_00DF2FCE: push eax
  loc_00DF2FCF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF2FD5: mov var_D0, eax
  loc_00DF2FDB: jmp 00DF2FE7h
  loc_00DF2FDD: mov var_D0, 00000000h
  loc_00DF2FE7: mov ecx, Me
  loc_00DF2FEA: mov dx, var_48
  loc_00DF2FEE: and dx, [ecx+0000006Eh]
  loc_00DF2FF2: mov var_6C, dx
  loc_00DF2FF6: lea ecx, var_44
  loc_00DF2FF9: call [00401254h] ; __vbaFreeObj
  loc_00DF2FFF: movsx eax, var_6C
  loc_00DF3003: test eax, eax
  loc_00DF3005: jz 00DF311Ah
  loc_00DF300B: mov var_4, 0000003Eh
  loc_00DF3012: mov ecx, Me
  loc_00DF3015: cmp [ecx+00000044h], 0000000Bh
  loc_00DF3019: jz 00DF3024h
  loc_00DF301B: mov edx, Me
  loc_00DF301E: cmp [edx+00000044h], 00000000h
  loc_00DF3022: jnz 00DF3066h
  loc_00DF3024: mov var_4, 0000003Fh
  loc_00DF302B: mov eax, Me
  loc_00DF302E: mov ecx, [eax+000001D0h]
  loc_00DF3034: sub ecx, 00000004h
  loc_00DF3037: jo 00DF32DCh
  loc_00DF303D: push ecx
  loc_00DF303E: mov edx, Me
  loc_00DF3041: mov eax, [edx+000001D4h]
  loc_00DF3047: sub eax, 00000004h
  loc_00DF304A: jo 00DF32DCh
  loc_00DF3050: push eax
  loc_00DF3051: push 00000004h
  loc_00DF3053: push 00000004h
  loc_00DF3055: lea ecx, var_30
  loc_00DF3058: push ecx
  loc_00DF3059: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF305E: call [00401074h] ; __vbaSetSystemError
  loc_00DF3064: jmp 00DF30A6h
  loc_00DF3066: mov var_4, 00000041h
  loc_00DF306D: mov edx, Me
  loc_00DF3070: mov eax, [edx+000001D0h]
  loc_00DF3076: sub eax, 00000003h
  loc_00DF3079: jo 00DF32DCh
  loc_00DF307F: push eax
  loc_00DF3080: mov ecx, Me
  loc_00DF3083: mov edx, [ecx+000001D4h]
  loc_00DF3089: sub edx, 00000003h
  loc_00DF308C: jo 00DF32DCh
  loc_00DF3092: push edx
  loc_00DF3093: push 00000003h
  loc_00DF3095: push 00000003h
  loc_00DF3097: lea eax, var_30
  loc_00DF309A: push eax
  loc_00DF309B: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF30A0: call [00401074h] ; __vbaSetSystemError
  loc_00DF30A6: mov var_4, 00000043h
  loc_00DF30AD: mov ecx, Me
  loc_00DF30B0: movsx edx, [ecx+00000070h]
  loc_00DF30B4: test edx, edx
  loc_00DF30B6: jz 00DF311Ah
  loc_00DF30B8: mov var_4, 00000044h
  loc_00DF30BF: lea eax, var_4C
  loc_00DF30C2: push eax
  loc_00DF30C3: mov ecx, Me
  loc_00DF30C6: mov edx, [ecx]
  loc_00DF30C8: mov eax, Me
  loc_00DF30CB: push eax
  loc_00DF30CC: call [edx+000000D8h]
  loc_00DF30D2: fnclex
  loc_00DF30D4: mov var_60, eax
  loc_00DF30D7: cmp var_60, 00000000h
  loc_00DF30DB: jge 00DF30FDh
  loc_00DF30DD: push 000000D8h
  loc_00DF30E2: push 006DCB00h
  loc_00DF30E7: mov ecx, Me
  loc_00DF30EA: push ecx
  loc_00DF30EB: mov edx, var_60
  loc_00DF30EE: push edx
  loc_00DF30EF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF30F5: mov var_D4, eax
  loc_00DF30FB: jmp 00DF3107h
  loc_00DF30FD: mov var_D4, 00000000h
  loc_00DF3107: lea eax, var_30
  loc_00DF310A: push eax
  loc_00DF310B: mov ecx, var_4C
  loc_00DF310E: push ecx
  loc_00DF310F: call 006DDA18h ; DrawFocusRect(%x1v, %x2v)
  loc_00DF3114: call [00401074h] ; __vbaSetSystemError
  loc_00DF311A: mov var_4, 00000047h
  loc_00DF3121: cmp arg_C, 00000002h
  loc_00DF3125: jz 00DF3297h
  loc_00DF312B: mov edx, Me
  loc_00DF312E: cmp [edx+00000044h], 00000000h
  loc_00DF3132: jnz 00DF3297h
  loc_00DF3138: mov var_4, 00000048h
  loc_00DF313F: mov eax, Me
  loc_00DF3142: mov ecx, [eax+000001D0h]
  loc_00DF3148: sub ecx, 00000001h
  loc_00DF314B: jo 00DF32DCh
  loc_00DF3151: push ecx
  loc_00DF3152: mov edx, Me
  loc_00DF3155: mov eax, [edx+000001D4h]
  loc_00DF315B: sub eax, 00000001h
  loc_00DF315E: jo 00DF32DCh
  loc_00DF3164: push eax
  loc_00DF3165: push 00000000h
  loc_00DF3167: push 00000000h
  loc_00DF3169: lea ecx, var_40
  loc_00DF316C: push ecx
  loc_00DF316D: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF3172: call [00401074h] ; __vbaSetSystemError
  loc_00DF3178: mov var_4, 00000049h
  loc_00DF317F: lea edx, var_4C
  loc_00DF3182: push edx
  loc_00DF3183: mov eax, Me
  loc_00DF3186: mov ecx, [eax]
  loc_00DF3188: mov edx, Me
  loc_00DF318B: push edx
  loc_00DF318C: call [ecx+000000D8h]
  loc_00DF3192: fnclex
  loc_00DF3194: mov var_60, eax
  loc_00DF3197: cmp var_60, 00000000h
  loc_00DF319B: jge 00DF31BDh
  loc_00DF319D: push 000000D8h
  loc_00DF31A2: push 006DCB00h
  loc_00DF31A7: mov eax, Me
  loc_00DF31AA: push eax
  loc_00DF31AB: mov ecx, var_60
  loc_00DF31AE: push ecx
  loc_00DF31AF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF31B5: mov var_D8, eax
  loc_00DF31BB: jmp 00DF31C7h
  loc_00DF31BD: mov var_D8, 00000000h
  loc_00DF31C7: push 0000000Fh
  loc_00DF31C9: push 00000005h
  loc_00DF31CB: lea edx, var_40
  loc_00DF31CE: push edx
  loc_00DF31CF: mov eax, var_4C
  loc_00DF31D2: push eax
  loc_00DF31D3: call 006DD9D0h ; DrawEdge(%x1v, %x2v, %x3v, %x4v)
  loc_00DF31D8: call [00401074h] ; __vbaSetSystemError
  loc_00DF31DE: mov var_4, 0000004Ah
  loc_00DF31E5: mov var_4C, 00000000h
  loc_00DF31EC: lea ecx, var_50
  loc_00DF31EF: push ecx
  loc_00DF31F0: lea edx, var_4C
  loc_00DF31F3: push edx
  loc_00DF31F4: push 8000000Ch
  loc_00DF31F9: mov eax, Me
  loc_00DF31FC: mov ecx, [eax]
  loc_00DF31FE: mov edx, Me
  loc_00DF3201: push edx
  loc_00DF3202: call [ecx+0000090Ch]
  loc_00DF3208: mov eax, var_50
  loc_00DF320B: push eax
  loc_00DF320C: mov ecx, Me
  loc_00DF320F: mov edx, [ecx+000001D0h]
  loc_00DF3215: sub edx, 00000001h
  loc_00DF3218: jo 00DF32DCh
  loc_00DF321E: push edx
  loc_00DF321F: mov eax, Me
  loc_00DF3222: mov ecx, [eax+000001D4h]
  loc_00DF3228: sub ecx, 00000001h
  loc_00DF322B: jo 00DF32DCh
  loc_00DF3231: push ecx
  loc_00DF3232: push 00000000h
  loc_00DF3234: push 00000000h
  loc_00DF3236: mov edx, Me
  loc_00DF3239: mov eax, [edx]
  loc_00DF323B: mov ecx, Me
  loc_00DF323E: push ecx
  loc_00DF323F: call [eax+000008F0h]
  loc_00DF3245: mov var_4, 0000004Bh
  loc_00DF324C: mov var_4C, 00000000h
  loc_00DF3253: lea edx, var_50
  loc_00DF3256: push edx
  loc_00DF3257: lea eax, var_4C
  loc_00DF325A: push eax
  loc_00DF325B: push 00000000h
  loc_00DF325D: mov ecx, Me
  loc_00DF3260: mov edx, [ecx]
  loc_00DF3262: mov eax, Me
  loc_00DF3265: push eax
  loc_00DF3266: call [edx+0000090Ch]
  loc_00DF326C: mov ecx, var_50
  loc_00DF326F: push ecx
  loc_00DF3270: mov edx, Me
  loc_00DF3273: mov eax, [edx+000001D0h]
  loc_00DF3279: push eax
  loc_00DF327A: mov ecx, Me
  loc_00DF327D: mov edx, [ecx+000001D4h]
  loc_00DF3283: push edx
  loc_00DF3284: push 00000000h
  loc_00DF3286: push 00000000h
  loc_00DF3288: mov eax, Me
  loc_00DF328B: mov ecx, [eax]
  loc_00DF328D: mov edx, Me
  loc_00DF3290: push edx
  loc_00DF3291: call [ecx+000008F0h]
  loc_00DF3297: fwait
  loc_00DF3298: push 00DF32AAh
  loc_00DF329D: jmp 00DF32A9h
  loc_00DF329F: lea ecx, var_44
  loc_00DF32A2: call [00401254h] ; __vbaFreeObj
  loc_00DF32A8: ret
  loc_00DF32A9: ret
  loc_00DF32AA: xor eax, eax
  loc_00DF32AC: mov ecx, var_20
  loc_00DF32AF: mov fs:[00000000h], ecx
  loc_00DF32B6: pop edi
  loc_00DF32B7: pop esi
  loc_00DF32B8: pop ebx
  loc_00DF32B9: mov esp, ebp
  loc_00DF32BB: pop ebp
  loc_00DF32BC: retn 0008h
End Sub

Private Sub Proc_2_125_DF32F0(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28) 'DF32F0
  loc_00DF32F0: sub esp, 00000030h
  loc_00DF32F3: push ebx
  loc_00DF32F4: xor eax, eax
  loc_00DF32F6: push ebp
  loc_00DF32F7: push esi
  loc_00DF32F8: mov esi, arg_38
  loc_00DF32FC: mov arg_24, eax
  loc_00DF3300: mov arg_28, eax
  loc_00DF3304: lea edx, Proc_2_125_DF32F0
  loc_00DF3308: mov ecx, [esi]
  loc_00DF330A: push edi
  loc_00DF330B: mov arg_30, eax
  loc_00DF330F: xor edi, edi
  loc_00DF3311: push edx
  loc_00DF3312: push esi
  loc_00DF3313: mov arg_3C, eax
  loc_00DF3317: mov arg_10, edi
  loc_00DF331B: mov arg_14, edi
  loc_00DF331F: mov arg_1C, edi
  loc_00DF3323: mov arg_18, edi
  loc_00DF3327: mov arg_20, edi
  loc_00DF332B: mov arg_24, edi
  loc_00DF332F: call [ecx+00000108h]
  loc_00DF3335: mov ebp, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF333B: cmp eax, edi
  loc_00DF333D: fnclex
  loc_00DF333F: jge 00DF334Fh
  loc_00DF3341: push 00000108h
  loc_00DF3346: push 006DCB00h
  loc_00DF334B: push esi
  loc_00DF334C: push eax
  loc_00DF334D: call ebp
  loc_00DF334F: fld real4 ptr Me
  loc_00DF3353: mov ebx, [00401208h] ; __vbaFpI4
  loc_00DF3359: call ebx
  loc_00DF335B: lea ecx, Me
  loc_00DF335F: mov [esi+000001D0h], eax
  loc_00DF3365: mov eax, [esi]
  loc_00DF3367: push ecx
  loc_00DF3368: push esi
  loc_00DF3369: call [eax+00000100h]
  loc_00DF336F: cmp eax, edi
  loc_00DF3371: fnclex
  loc_00DF3373: jge 00DF3383h
  loc_00DF3375: push 00000100h
  loc_00DF337A: push 006DCB00h
  loc_00DF337F: push esi
  loc_00DF3380: push eax
  loc_00DF3381: call ebp
  loc_00DF3383: fld real4 ptr Me
  loc_00DF3387: call ebx
  loc_00DF3389: mov edx, [esi]
  loc_00DF338B: mov [esi+000001D4h], eax
  loc_00DF3391: lea eax, arg_C
  loc_00DF3395: lea ecx, Me
  loc_00DF3399: push eax
  loc_00DF339A: mov eax, [esi+00000088h]
  loc_00DF33A0: push ecx
  loc_00DF33A1: push eax
  loc_00DF33A2: push esi
  loc_00DF33A3: mov arg_18, edi
  loc_00DF33A7: call [edx+0000090Ch]
  loc_00DF33AD: mov Me, edi
  loc_00DF33B1: mov ebp, Proc_2_125_DF32F00
  loc_00DF33B5: mov ecx, arg_C
  loc_00DF33B9: cmp ebp, 00000002h
  loc_00DF33BC: mov arg_3C, ecx
  loc_00DF33C0: jnz 00DF33E6h
  loc_00DF33C2: mov edx, [esi]
  loc_00DF33C4: lea eax, arg_C
  loc_00DF33C8: lea ecx, Me
  loc_00DF33CC: push eax
  loc_00DF33CD: push ecx
  loc_00DF33CE: push 00FFFFFFh
  loc_00DF33D3: push esi
  loc_00DF33D4: call [edx+0000090Ch]
  loc_00DF33DA: mov edx, arg_C
  loc_00DF33DE: mov [esi+00000090h], edx
  loc_00DF33E4: jmp 00DF3408h
  loc_00DF33E6: mov eax, [esi]
  loc_00DF33E8: lea ecx, arg_C
  loc_00DF33EC: lea edx, Me
  loc_00DF33F0: push ecx
  loc_00DF33F1: push edx
  loc_00DF33F2: push 80000012h
  loc_00DF33F7: push esi
  loc_00DF33F8: call [eax+0000090Ch]
  loc_00DF33FE: mov eax, arg_C
  loc_00DF3402: mov [esi+00000090h], eax
  loc_00DF3408: mov eax, [esi+00000064h]
  loc_00DF340B: cmp eax, edi
  loc_00DF340D: jz 00DF36CEh
  loc_00DF3413: cmp [esi+0000006Ch], di
  loc_00DF3417: jz 00DF3424h
  loc_00DF3419: cmp [esi+0000004Ch], di
  loc_00DF341D: jz 00DF3424h
  loc_00DF341F: mov ebp, 00000002h
  loc_00DF3424: cmp eax, edi
  loc_00DF3426: jz 00DF36CEh
  loc_00DF342C: cmp [esi+0000006Ch], di
  loc_00DF3430: jz 00DF36CEh
  loc_00DF3436: cmp ebp, 00000002h
  loc_00DF3439: jz 00DF36CEh
  loc_00DF343F: mov ecx, [esi+000001D0h]
  loc_00DF3445: mov edx, [esi+000001D4h]
  loc_00DF344B: push ecx
  loc_00DF344C: push edx
  loc_00DF344D: push edi
  loc_00DF344E: lea eax, arg_34
  loc_00DF3452: push edi
  loc_00DF3453: push eax
  loc_00DF3454: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF3459: call [00401074h] ; __vbaSetSystemError
  loc_00DF345F: mov ecx, [esi]
  loc_00DF3461: lea edx, arg_C
  loc_00DF3465: lea eax, Me
  loc_00DF3469: push edx
  loc_00DF346A: push eax
  loc_00DF346B: push 00FEFEFEh
  loc_00DF3470: push esi
  loc_00DF3471: mov arg_18, edi
  loc_00DF3475: call [ecx+0000090Ch]
  loc_00DF347B: mov ecx, [esi]
  loc_00DF347D: lea edx, arg_28
  loc_00DF3481: mov eax, arg_C
  loc_00DF3485: push edx
  loc_00DF3486: push eax
  loc_00DF3487: push esi
  loc_00DF3488: call [ecx+00000974h]
  loc_00DF348E: mov ecx, [esi]
  loc_00DF3490: lea edx, arg_C
  loc_00DF3494: lea eax, Me
  loc_00DF3498: push edx
  loc_00DF3499: push eax
  loc_00DF349A: push 80000012h
  loc_00DF349F: push esi
  loc_00DF34A0: mov arg_18, edi
  loc_00DF34A4: call [ecx+0000090Ch]
  loc_00DF34AA: mov edx, [esi]
  loc_00DF34AC: push esi
  loc_00DF34AD: mov ecx, arg_10
  loc_00DF34B1: mov [esi+00000090h], ecx
  loc_00DF34B7: call [edx+00000920h]
  loc_00DF34BD: mov eax, [esi]
  loc_00DF34BF: lea ecx, arg_C
  loc_00DF34C3: lea edx, Me
  loc_00DF34C7: push ecx
  loc_00DF34C8: push edx
  loc_00DF34C9: push 00AF987Ah
  loc_00DF34CE: push esi
  loc_00DF34CF: mov arg_18, edi
  loc_00DF34D3: call [eax+0000090Ch]
  loc_00DF34D9: mov edx, [esi+000001D0h]
  loc_00DF34DF: mov eax, [esi]
  loc_00DF34E1: mov ecx, arg_C
  loc_00DF34E5: push ecx
  loc_00DF34E6: mov ecx, [esi+000001D4h]
  loc_00DF34EC: push edx
  loc_00DF34ED: push ecx
  loc_00DF34EE: push edi
  loc_00DF34EF: push edi
  loc_00DF34F0: push esi
  loc_00DF34F1: call [eax+000008F0h]
  loc_00DF34F7: mov edx, [esi]
  loc_00DF34F9: lea eax, arg_C
  loc_00DF34FD: lea ecx, Me
  loc_00DF3501: push eax
  loc_00DF3502: push ecx
  loc_00DF3503: push 00C1B3A0h
  loc_00DF3508: push esi
  loc_00DF3509: mov arg_18, edi
  loc_00DF350D: call [edx+0000090Ch]
  loc_00DF3513: fld real8 ptr [004015A0h]
  loc_00DF3519: mov edx, arg_C
  loc_00DF351D: lea ecx, arg_18
  loc_00DF3521: fstp real4 ptr arg_10
  loc_00DF3525: fnstsw ax
  loc_00DF3527: test al, 0Dh
  loc_00DF3529: jnz 00DF3D48h
  loc_00DF352F: mov arg_14, edx
  loc_00DF3533: lea edx, arg_10
  loc_00DF3537: mov eax, [esi]
  loc_00DF3539: push ecx
  loc_00DF353A: push edx
  loc_00DF353B: lea ecx, arg_1C
  loc_00DF353F: push ecx
  loc_00DF3540: push esi
  loc_00DF3541: call [eax+0000097Ch]
  loc_00DF3547: mov eax, [esi]
  loc_00DF3549: lea ecx, arg_1C
  loc_00DF354D: mov edx, arg_18
  loc_00DF3551: push ecx
  loc_00DF3552: push esi
  loc_00DF3553: mov arg_24, edx
  loc_00DF3557: call [eax+00000940h]
  loc_00DF355D: cmp ebp, 00000001h
  loc_00DF3560: jnz 00DF3D3Ch
  loc_00DF3566: mov edx, [esi]
  loc_00DF3568: lea eax, arg_C
  loc_00DF356C: lea ecx, Me
  loc_00DF3570: push eax
  loc_00DF3571: push ecx
  loc_00DF3572: push 00EDF0F2h
  loc_00DF3577: push esi
  loc_00DF3578: mov arg_18, edi
  loc_00DF357C: call [edx+0000090Ch]
  loc_00DF3582: mov eax, [esi+000001D4h]
  loc_00DF3588: mov edx, [esi]
  loc_00DF358A: mov ecx, arg_C
  loc_00DF358E: push ecx
  loc_00DF358F: mov ecx, [esi+000001D0h]
  loc_00DF3595: sub ecx, 00000002h
  loc_00DF3598: jo 00DF3D4Dh
  loc_00DF359E: push ecx
  loc_00DF359F: mov ecx, eax
  loc_00DF35A1: sub ecx, 00000002h
  loc_00DF35A4: jo 00DF3D4Dh
  loc_00DF35AA: sub eax, 00000002h
  loc_00DF35AD: push ecx
  loc_00DF35AE: push 00000002h
  loc_00DF35B0: jo 00DF3D4Dh
  loc_00DF35B6: push eax
  loc_00DF35B7: push esi
  loc_00DF35B8: call [edx+000008E4h]
  loc_00DF35BE: mov edx, [esi]
  loc_00DF35C0: lea eax, arg_C
  loc_00DF35C4: lea ecx, Me
  loc_00DF35C8: push eax
  loc_00DF35C9: push ecx
  loc_00DF35CA: push 00D8DEE4h
  loc_00DF35CF: push esi
  loc_00DF35D0: mov arg_18, edi
  loc_00DF35D4: call [edx+0000090Ch]
  loc_00DF35DA: mov eax, [esi+000001D0h]
  loc_00DF35E0: mov edx, [esi]
  loc_00DF35E2: mov ecx, arg_C
  loc_00DF35E6: push ecx
  loc_00DF35E7: mov ecx, eax
  loc_00DF35E9: sub ecx, 00000002h
  loc_00DF35EC: jo 00DF3D4Dh
  loc_00DF35F2: push ecx
  loc_00DF35F3: mov ecx, [esi+000001D4h]
  loc_00DF35F9: sub ecx, 00000002h
  loc_00DF35FC: jo 00DF3D4Dh
  loc_00DF3602: sub eax, 00000002h
  loc_00DF3605: push ecx
  loc_00DF3606: jo 00DF3D4Dh
  loc_00DF360C: push eax
  loc_00DF360D: push 00000002h
  loc_00DF360F: push esi
  loc_00DF3610: call [edx+000008E4h]
  loc_00DF3616: mov edx, [esi]
  loc_00DF3618: lea eax, arg_C
  loc_00DF361C: lea ecx, Me
  loc_00DF3620: push eax
  loc_00DF3621: push ecx
  loc_00DF3622: push 00E8ECEFh
  loc_00DF3627: push esi
  loc_00DF3628: mov arg_18, edi
  loc_00DF362C: call [edx+0000090Ch]
  loc_00DF3632: mov eax, [esi+000001D0h]
  loc_00DF3638: mov edx, [esi]
  loc_00DF363A: mov ecx, arg_C
  loc_00DF363E: push ecx
  loc_00DF363F: mov ecx, eax
  loc_00DF3641: sub ecx, 00000003h
  loc_00DF3644: jo 00DF3D4Dh
  loc_00DF364A: push ecx
  loc_00DF364B: mov ecx, [esi+000001D4h]
  loc_00DF3651: sub ecx, ebp
  loc_00DF3653: jo 00DF3D4Dh
  loc_00DF3659: sub eax, 00000003h
  loc_00DF365C: push ecx
  loc_00DF365D: jo 00DF3D4Dh
  loc_00DF3663: push eax
  loc_00DF3664: push ebp
  loc_00DF3665: push esi
  loc_00DF3666: call [edx+000008E4h]
  loc_00DF366C: mov edx, [esi]
  loc_00DF366E: lea eax, arg_C
  loc_00DF3672: lea ecx, Me
  loc_00DF3676: push eax
  loc_00DF3677: push ecx
  loc_00DF3678: push 00F8F9FAh
  loc_00DF367D: push esi
  loc_00DF367E: mov arg_18, edi
  loc_00DF3682: call [edx+0000090Ch]
  loc_00DF3688: mov eax, [esi+000001D0h]
  loc_00DF368E: mov edx, [esi]
  loc_00DF3690: mov ecx, arg_C
  loc_00DF3694: push ecx
  loc_00DF3695: mov ecx, eax
  loc_00DF3697: sub ecx, 00000004h
  loc_00DF369A: jo 00DF3D4Dh
  loc_00DF36A0: push ecx
  loc_00DF36A1: mov ecx, [esi+000001D4h]
  loc_00DF36A7: sub ecx, ebp
  loc_00DF36A9: jo 00DF3D4Dh
  loc_00DF36AF: sub eax, 00000004h
  loc_00DF36B2: push ecx
  loc_00DF36B3: jo 00DF3D4Dh
  loc_00DF36B9: push eax
  loc_00DF36BA: push ebp
  loc_00DF36BB: push esi
  loc_00DF36BC: call [edx+000008E4h]
  loc_00DF36C2: pop edi
  loc_00DF36C3: pop esi
  loc_00DF36C4: pop ebp
  loc_00DF36C5: xor eax, eax
  loc_00DF36C7: pop ebx
  loc_00DF36C8: add esp, 00000030h
  loc_00DF36CB: retn 0008h
End Sub

Private Sub Proc_2_126_DF3D60(arg_C) 'DF3D60
  loc_00DF3D60: push ebp
  loc_00DF3D61: mov ebp, esp
  loc_00DF3D63: sub esp, 00000018h
  loc_00DF3D66: push 00402836h ; __vbaExceptHandler
  loc_00DF3D6B: mov eax, fs:[00000000h]
  loc_00DF3D71: push eax
  loc_00DF3D72: mov fs:[00000000h], esp
  loc_00DF3D79: mov eax, 000000B0h
  loc_00DF3D7E: call 00402830h ; __vbaChkstk
  loc_00DF3D83: push ebx
  loc_00DF3D84: push esi
  loc_00DF3D85: push edi
  loc_00DF3D86: mov var_18, esp
  loc_00DF3D89: mov var_14, 004015A8h
  loc_00DF3D90: mov var_10, 00000000h
  loc_00DF3D97: mov var_C, 00000000h
  loc_00DF3D9E: mov var_4, 00000001h
  loc_00DF3DA5: mov var_4, 00000002h
  loc_00DF3DAC: lea eax, var_38
  loc_00DF3DAF: push eax
  loc_00DF3DB0: mov ecx, Me
  loc_00DF3DB3: mov edx, [ecx]
  loc_00DF3DB5: mov eax, Me
  loc_00DF3DB8: push eax
  loc_00DF3DB9: call [edx+00000108h]
  loc_00DF3DBF: fnclex
  loc_00DF3DC1: mov var_50, eax
  loc_00DF3DC4: cmp var_50, 00000000h
  loc_00DF3DC8: jge 00DF3DE7h
  loc_00DF3DCA: push 00000108h
  loc_00DF3DCF: push 006DCB00h
  loc_00DF3DD4: mov ecx, Me
  loc_00DF3DD7: push ecx
  loc_00DF3DD8: mov edx, var_50
  loc_00DF3DDB: push edx
  loc_00DF3DDC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF3DE2: mov var_80, eax
  loc_00DF3DE5: jmp 00DF3DEEh
  loc_00DF3DE7: mov var_80, 00000000h
  loc_00DF3DEE: fld real4 ptr var_38
  loc_00DF3DF1: call [00401208h] ; __vbaFpI4
  loc_00DF3DF7: mov ecx, Me
  loc_00DF3DFA: mov [ecx+000001D0h], eax
  loc_00DF3E00: mov var_4, 00000003h
  loc_00DF3E07: lea edx, var_38
  loc_00DF3E0A: push edx
  loc_00DF3E0B: mov eax, Me
  loc_00DF3E0E: mov ecx, [eax]
  loc_00DF3E10: mov edx, Me
  loc_00DF3E13: push edx
  loc_00DF3E14: call [ecx+00000100h]
  loc_00DF3E1A: fnclex
  loc_00DF3E1C: mov var_50, eax
  loc_00DF3E1F: cmp var_50, 00000000h
  loc_00DF3E23: jge 00DF3E45h
  loc_00DF3E25: push 00000100h
  loc_00DF3E2A: push 006DCB00h
  loc_00DF3E2F: mov eax, Me
  loc_00DF3E32: push eax
  loc_00DF3E33: mov ecx, var_50
  loc_00DF3E36: push ecx
  loc_00DF3E37: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF3E3D: mov var_84, eax
  loc_00DF3E43: jmp 00DF3E4Fh
  loc_00DF3E45: mov var_84, 00000000h
  loc_00DF3E4F: fld real4 ptr var_38
  loc_00DF3E52: call [00401208h] ; __vbaFpI4
  loc_00DF3E58: mov edx, Me
  loc_00DF3E5B: mov [edx+000001D4h], eax
  loc_00DF3E61: mov var_4, 00000004h
  loc_00DF3E68: mov var_38, 00000000h
  loc_00DF3E6F: lea eax, var_3C
  loc_00DF3E72: push eax
  loc_00DF3E73: lea ecx, var_38
  loc_00DF3E76: push ecx
  loc_00DF3E77: mov edx, Me
  loc_00DF3E7A: mov eax, [edx+00000088h]
  loc_00DF3E80: push eax
  loc_00DF3E81: mov ecx, Me
  loc_00DF3E84: mov edx, [ecx]
  loc_00DF3E86: mov eax, Me
  loc_00DF3E89: push eax
  loc_00DF3E8A: call [edx+0000090Ch]
  loc_00DF3E90: mov ecx, var_3C
  loc_00DF3E93: mov var_34, ecx
  loc_00DF3E96: mov var_4, 00000005h
  loc_00DF3E9D: mov edx, Me
  loc_00DF3EA0: mov eax, [edx+000001D0h]
  loc_00DF3EA6: push eax
  loc_00DF3EA7: mov ecx, Me
  loc_00DF3EAA: mov edx, [ecx+000001D4h]
  loc_00DF3EB0: push edx
  loc_00DF3EB1: push 00000000h
  loc_00DF3EB3: push 00000000h
  loc_00DF3EB5: mov eax, Me
  loc_00DF3EB8: add eax, 000000ACh
  loc_00DF3EBD: push eax
  loc_00DF3EBE: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF3EC3: call [00401074h] ; __vbaSetSystemError
  loc_00DF3EC9: mov var_4, 00000006h
  loc_00DF3ED0: mov ecx, Me
  loc_00DF3ED3: movsx edx, [ecx+0000007Ch]
  loc_00DF3ED7: test edx, edx
  loc_00DF3ED9: jnz 00DF4032h
  loc_00DF3EDF: mov var_4, 00000007h
  loc_00DF3EE6: mov eax, Me
  loc_00DF3EE9: mov ecx, [eax]
  loc_00DF3EEB: mov edx, Me
  loc_00DF3EEE: push edx
  loc_00DF3EEF: call [ecx+00000914h]
  loc_00DF3EF5: mov var_4, 00000008h
  loc_00DF3EFC: push 0000000Fh
  loc_00DF3EFE: call 006DDDB4h ; GetSysColor(%x1v)
  loc_00DF3F03: mov var_38, eax
  loc_00DF3F06: call [00401074h] ; __vbaSetSystemError
  loc_00DF3F0C: mov var_3C, 3DCCCCCDh
  loc_00DF3F13: lea eax, var_40
  loc_00DF3F16: push eax
  loc_00DF3F17: lea ecx, var_3C
  loc_00DF3F1A: push ecx
  loc_00DF3F1B: lea edx, var_34
  loc_00DF3F1E: push edx
  loc_00DF3F1F: mov eax, Me
  loc_00DF3F22: mov ecx, [eax]
  loc_00DF3F24: mov edx, Me
  loc_00DF3F27: push edx
  loc_00DF3F28: call [ecx+0000097Ch]
  loc_00DF3F2E: lea eax, var_44
  loc_00DF3F31: push eax
  loc_00DF3F32: mov ecx, var_40
  loc_00DF3F35: push ecx
  loc_00DF3F36: mov edx, var_38
  loc_00DF3F39: push edx
  loc_00DF3F3A: mov eax, Me
  loc_00DF3F3D: mov ecx, [eax]
  loc_00DF3F3F: mov edx, Me
  loc_00DF3F42: push edx
  loc_00DF3F43: call [ecx+000008ECh]
  loc_00DF3F49: mov eax, Me
  loc_00DF3F4C: add eax, 000000ACh
  loc_00DF3F51: push eax
  loc_00DF3F52: mov ecx, var_44
  loc_00DF3F55: push ecx
  loc_00DF3F56: mov edx, Me
  loc_00DF3F59: mov eax, [edx]
  loc_00DF3F5B: mov ecx, Me
  loc_00DF3F5E: push ecx
  loc_00DF3F5F: call [eax+00000974h]
  loc_00DF3F65: mov var_4, 00000009h
  loc_00DF3F6C: mov edx, Me
  loc_00DF3F6F: mov eax, [edx]
  loc_00DF3F71: mov ecx, Me
  loc_00DF3F74: push ecx
  loc_00DF3F75: call [eax+00000920h]
  loc_00DF3F7B: mov var_4, 0000000Ah
  loc_00DF3F82: fld real8 ptr [004017E0h]
  loc_00DF3F88: fchs
  loc_00DF3F8A: fstp real4 ptr var_38
  loc_00DF3F8D: fnstsw ax
  loc_00DF3F8F: test al, 0Dh
  loc_00DF3F91: jnz 00DF5EFBh
  loc_00DF3F97: lea edx, var_3C
  loc_00DF3F9A: push edx
  loc_00DF3F9B: lea eax, var_38
  loc_00DF3F9E: push eax
  loc_00DF3F9F: lea ecx, var_34
  loc_00DF3FA2: push ecx
  loc_00DF3FA3: mov edx, Me
  loc_00DF3FA6: mov eax, [edx]
  loc_00DF3FA8: mov ecx, Me
  loc_00DF3FAB: push ecx
  loc_00DF3FAC: call [eax+0000097Ch]
  loc_00DF3FB2: mov edx, var_3C
  loc_00DF3FB5: push edx
  loc_00DF3FB6: mov eax, Me
  loc_00DF3FB9: mov ecx, [eax+000001D0h]
  loc_00DF3FBF: push ecx
  loc_00DF3FC0: mov edx, Me
  loc_00DF3FC3: mov eax, [edx+000001D4h]
  loc_00DF3FC9: push eax
  loc_00DF3FCA: push 00000000h
  loc_00DF3FCC: push 00000000h
  loc_00DF3FCE: mov ecx, Me
  loc_00DF3FD1: mov edx, [ecx]
  loc_00DF3FD3: mov eax, Me
  loc_00DF3FD6: push eax
  loc_00DF3FD7: call [edx+000008F0h]
  loc_00DF3FDD: mov var_4, 0000000Bh
  loc_00DF3FE4: fld real8 ptr [004017E0h]
  loc_00DF3FEA: fchs
  loc_00DF3FEC: fstp real4 ptr var_38
  loc_00DF3FEF: fnstsw ax
  loc_00DF3FF1: test al, 0Dh
  loc_00DF3FF3: jnz 00DF5EFBh
  loc_00DF3FF9: lea ecx, var_3C
  loc_00DF3FFC: push ecx
  loc_00DF3FFD: lea edx, var_38
  loc_00DF4000: push edx
  loc_00DF4001: lea eax, var_34
  loc_00DF4004: push eax
  loc_00DF4005: mov ecx, Me
  loc_00DF4008: mov edx, [ecx]
  loc_00DF400A: mov eax, Me
  loc_00DF400D: push eax
  loc_00DF400E: call [edx+0000097Ch]
  loc_00DF4014: mov ecx, var_3C
  loc_00DF4017: mov var_40, ecx
  loc_00DF401A: lea edx, var_40
  loc_00DF401D: push edx
  loc_00DF401E: mov eax, Me
  loc_00DF4021: mov ecx, [eax]
  loc_00DF4023: mov edx, Me
  loc_00DF4026: push edx
  loc_00DF4027: call [ecx+00000940h]
  loc_00DF402D: jmp 00DF5E96h
  loc_00DF4032: mov var_4, 0000000Eh
  loc_00DF4039: mov eax, arg_C
  loc_00DF403C: mov var_54, eax
  loc_00DF403F: mov ecx, var_54
  loc_00DF4042: mov var_88, ecx
  loc_00DF4048: cmp var_88, 00000000h
  loc_00DF404F: jz 00DF4075h
  loc_00DF4051: cmp var_88, 00000001h
  loc_00DF4058: jz 00DF46D1h
  loc_00DF405E: cmp var_88, 00000002h
  loc_00DF4065: jz 00DF4CFCh
  loc_00DF406B: jmp 00DF51C6h
  loc_00DF4070: jmp 00DF51C6h
  loc_00DF4075: mov var_4, 00000010h
  loc_00DF407C: mov edx, Me
  loc_00DF407F: mov eax, [edx]
  loc_00DF4081: mov ecx, Me
  loc_00DF4084: push ecx
  loc_00DF4085: call [eax+00000914h]
  loc_00DF408B: mov var_4, 00000011h
  loc_00DF4092: mov edx, Me
  loc_00DF4095: mov eax, [edx+000000CCh]
  loc_00DF409B: mov var_58, eax
  loc_00DF409E: mov ecx, var_58
  loc_00DF40A1: mov var_8C, ecx
  loc_00DF40A7: cmp var_8C, 00000003h
  loc_00DF40AE: ja 00DF46CCh
  loc_00DF40B4: mov edx, var_8C
  loc_00DF40BA: jmp [edx*4+00DF5EABh]
  loc_00DF40C1: jmp 00DF46CCh
  loc_00DF40C6: mov var_4, 00000013h
  loc_00DF40CD: mov var_38, 3D8F5C29h
  loc_00DF40D4: lea eax, var_3C
  loc_00DF40D7: push eax
  loc_00DF40D8: lea ecx, var_38
  loc_00DF40DB: push ecx
  loc_00DF40DC: lea edx, var_34
  loc_00DF40DF: push edx
  loc_00DF40E0: mov eax, Me
  loc_00DF40E3: mov ecx, [eax]
  loc_00DF40E5: mov edx, Me
  loc_00DF40E8: push edx
  loc_00DF40E9: call [ecx+0000097Ch]
  loc_00DF40EF: push 00000001h
  loc_00DF40F1: mov eax, var_34
  loc_00DF40F4: push eax
  loc_00DF40F5: mov ecx, var_3C
  loc_00DF40F8: push ecx
  loc_00DF40F9: mov edx, Me
  loc_00DF40FC: mov eax, [edx+000001D0h]
  loc_00DF4102: push eax
  loc_00DF4103: mov ecx, Me
  loc_00DF4106: mov edx, [ecx+000001D4h]
  loc_00DF410C: push edx
  loc_00DF410D: push 00000000h
  loc_00DF410F: push 00000000h
  loc_00DF4111: mov eax, Me
  loc_00DF4114: mov ecx, [eax]
  loc_00DF4116: mov edx, Me
  loc_00DF4119: push edx
  loc_00DF411A: call [ecx+00000908h]
  loc_00DF4120: mov var_4, 00000014h
  loc_00DF4127: mov var_38, 3DCCCCCDh
  loc_00DF412E: lea eax, var_3C
  loc_00DF4131: push eax
  loc_00DF4132: lea ecx, var_38
  loc_00DF4135: push ecx
  loc_00DF4136: lea edx, var_34
  loc_00DF4139: push edx
  loc_00DF413A: mov eax, Me
  loc_00DF413D: mov ecx, [eax]
  loc_00DF413F: mov edx, Me
  loc_00DF4142: push edx
  loc_00DF4143: call [ecx+0000097Ch]
  loc_00DF4149: mov var_40, 3DA3D70Ah
  loc_00DF4150: lea eax, var_44
  loc_00DF4153: push eax
  loc_00DF4154: lea ecx, var_40
  loc_00DF4157: push ecx
  loc_00DF4158: lea edx, var_34
  loc_00DF415B: push edx
  loc_00DF415C: mov eax, Me
  loc_00DF415F: mov ecx, [eax]
  loc_00DF4161: mov edx, Me
  loc_00DF4164: push edx
  loc_00DF4165: call [ecx+0000097Ch]
  loc_00DF416B: push 00000001h
  loc_00DF416D: mov eax, var_44
  loc_00DF4170: push eax
  loc_00DF4171: mov ecx, var_3C
  loc_00DF4174: push ecx
  loc_00DF4175: push 00000004h
  loc_00DF4177: mov edx, Me
  loc_00DF417A: mov eax, [edx+000001D4h]
  loc_00DF4180: push eax
  loc_00DF4181: push 00000000h
  loc_00DF4183: push 00000000h
  loc_00DF4185: mov ecx, Me
  loc_00DF4188: mov edx, [ecx]
  loc_00DF418A: mov eax, Me
  loc_00DF418D: push eax
  loc_00DF418E: call [edx+00000908h]
  loc_00DF4194: mov var_4, 00000015h
  loc_00DF419B: mov ecx, Me
  loc_00DF419E: mov edx, [ecx]
  loc_00DF41A0: mov eax, Me
  loc_00DF41A3: push eax
  loc_00DF41A4: call [edx+00000920h]
  loc_00DF41AA: mov var_4, 00000016h
  loc_00DF41B1: fld real8 ptr [004017D8h]
  loc_00DF41B7: fchs
  loc_00DF41B9: fstp real4 ptr var_38
  loc_00DF41BC: fnstsw ax
  loc_00DF41BE: test al, 0Dh
  loc_00DF41C0: jnz 00DF5EFBh
  loc_00DF41C6: lea ecx, var_3C
  loc_00DF41C9: push ecx
  loc_00DF41CA: lea edx, var_38
  loc_00DF41CD: push edx
  loc_00DF41CE: lea eax, var_34
  loc_00DF41D1: push eax
  loc_00DF41D2: mov ecx, Me
  loc_00DF41D5: mov edx, [ecx]
  loc_00DF41D7: mov eax, Me
  loc_00DF41DA: push eax
  loc_00DF41DB: call [edx+0000097Ch]
  loc_00DF41E1: mov ecx, var_3C
  loc_00DF41E4: push ecx
  loc_00DF41E5: mov edx, Me
  loc_00DF41E8: mov eax, [edx+000001D0h]
  loc_00DF41EE: sub eax, 00000002h
  loc_00DF41F1: jo 00DF5F00h
  loc_00DF41F7: push eax
  loc_00DF41F8: mov ecx, Me
  loc_00DF41FB: mov edx, [ecx+000001D4h]
  loc_00DF4201: sub edx, 00000002h
  loc_00DF4204: jo 00DF5F00h
  loc_00DF420A: push edx
  loc_00DF420B: mov eax, Me
  loc_00DF420E: mov ecx, [eax+000001D0h]
  loc_00DF4214: sub ecx, 00000002h
  loc_00DF4217: jo 00DF5F00h
  loc_00DF421D: push ecx
  loc_00DF421E: push 00000001h
  loc_00DF4220: mov edx, Me
  loc_00DF4223: mov eax, [edx]
  loc_00DF4225: mov ecx, Me
  loc_00DF4228: push ecx
  loc_00DF4229: call [eax+000008E4h]
  loc_00DF422F: mov var_4, 00000017h
  loc_00DF4236: fld real8 ptr [004017D0h]
  loc_00DF423C: fchs
  loc_00DF423E: fstp real4 ptr var_38
  loc_00DF4241: fnstsw ax
  loc_00DF4243: test al, 0Dh
  loc_00DF4245: jnz 00DF5EFBh
  loc_00DF424B: lea edx, var_3C
  loc_00DF424E: push edx
  loc_00DF424F: lea eax, var_38
  loc_00DF4252: push eax
  loc_00DF4253: lea ecx, var_34
  loc_00DF4256: push ecx
  loc_00DF4257: mov edx, Me
  loc_00DF425A: mov eax, [edx]
  loc_00DF425C: mov ecx, Me
  loc_00DF425F: push ecx
  loc_00DF4260: call [eax+0000097Ch]
  loc_00DF4266: mov edx, var_3C
  loc_00DF4269: push edx
  loc_00DF426A: mov eax, Me
  loc_00DF426D: mov ecx, [eax+000001D0h]
  loc_00DF4273: sub ecx, 00000003h
  loc_00DF4276: jo 00DF5F00h
  loc_00DF427C: push ecx
  loc_00DF427D: mov edx, Me
  loc_00DF4280: mov eax, [edx+000001D4h]
  loc_00DF4286: sub eax, 00000002h
  loc_00DF4289: jo 00DF5F00h
  loc_00DF428F: push eax
  loc_00DF4290: mov ecx, Me
  loc_00DF4293: mov edx, [ecx+000001D0h]
  loc_00DF4299: sub edx, 00000003h
  loc_00DF429C: jo 00DF5F00h
  loc_00DF42A2: push edx
  loc_00DF42A3: push 00000001h
  loc_00DF42A5: mov eax, Me
  loc_00DF42A8: mov ecx, [eax]
  loc_00DF42AA: mov edx, Me
  loc_00DF42AD: push edx
  loc_00DF42AE: call [ecx+000008E4h]
  loc_00DF42B4: mov var_4, 00000018h
  loc_00DF42BB: fld real8 ptr [004017C8h]
  loc_00DF42C1: fchs
  loc_00DF42C3: fstp real4 ptr var_38
  loc_00DF42C6: fnstsw ax
  loc_00DF42C8: test al, 0Dh
  loc_00DF42CA: jnz 00DF5EFBh
  loc_00DF42D0: lea eax, var_3C
  loc_00DF42D3: push eax
  loc_00DF42D4: lea ecx, var_38
  loc_00DF42D7: push ecx
  loc_00DF42D8: lea edx, var_34
  loc_00DF42DB: push edx
  loc_00DF42DC: mov eax, Me
  loc_00DF42DF: mov ecx, [eax]
  loc_00DF42E1: mov edx, Me
  loc_00DF42E4: push edx
  loc_00DF42E5: call [ecx+0000097Ch]
  loc_00DF42EB: mov eax, var_3C
  loc_00DF42EE: push eax
  loc_00DF42EF: mov ecx, Me
  loc_00DF42F2: mov edx, [ecx+000001D0h]
  loc_00DF42F8: sub edx, 00000004h
  loc_00DF42FB: jo 00DF5F00h
  loc_00DF4301: push edx
  loc_00DF4302: mov eax, Me
  loc_00DF4305: mov ecx, [eax+000001D4h]
  loc_00DF430B: sub ecx, 00000002h
  loc_00DF430E: jo 00DF5F00h
  loc_00DF4314: push ecx
  loc_00DF4315: mov edx, Me
  loc_00DF4318: mov eax, [edx+000001D0h]
  loc_00DF431E: sub eax, 00000004h
  loc_00DF4321: jo 00DF5F00h
  loc_00DF4327: push eax
  loc_00DF4328: push 00000001h
  loc_00DF432A: mov ecx, Me
  loc_00DF432D: mov edx, [ecx]
  loc_00DF432F: mov eax, Me
  loc_00DF4332: push eax
  loc_00DF4333: call [edx+000008E4h]
  loc_00DF4339: mov var_4, 00000019h
  loc_00DF4340: fld real8 ptr [004017C0h]
  loc_00DF4346: fchs
  loc_00DF4348: fstp real4 ptr var_38
  loc_00DF434B: fnstsw ax
  loc_00DF434D: test al, 0Dh
  loc_00DF434F: jnz 00DF5EFBh
  loc_00DF4355: lea ecx, var_3C
  loc_00DF4358: push ecx
  loc_00DF4359: lea edx, var_38
  loc_00DF435C: push edx
  loc_00DF435D: lea eax, var_34
  loc_00DF4360: push eax
  loc_00DF4361: mov ecx, Me
  loc_00DF4364: mov edx, [ecx]
  loc_00DF4366: mov eax, Me
  loc_00DF4369: push eax
  loc_00DF436A: call [edx+0000097Ch]
  loc_00DF4370: mov ecx, var_3C
  loc_00DF4373: push ecx
  loc_00DF4374: mov edx, Me
  loc_00DF4377: mov eax, [edx+000001D0h]
  loc_00DF437D: sub eax, 00000002h
  loc_00DF4380: jo 00DF5F00h
  loc_00DF4386: push eax
  loc_00DF4387: mov ecx, Me
  loc_00DF438A: mov edx, [ecx+000001D4h]
  loc_00DF4390: sub edx, 00000002h
  loc_00DF4393: jo 00DF5F00h
  loc_00DF4399: push edx
  loc_00DF439A: push 00000002h
  loc_00DF439C: mov eax, Me
  loc_00DF439F: mov ecx, [eax+000001D4h]
  loc_00DF43A5: sub ecx, 00000002h
  loc_00DF43A8: jo 00DF5F00h
  loc_00DF43AE: push ecx
  loc_00DF43AF: mov edx, Me
  loc_00DF43B2: mov eax, [edx]
  loc_00DF43B4: mov ecx, Me
  loc_00DF43B7: push ecx
  loc_00DF43B8: call [eax+000008E4h]
  loc_00DF43BE: mov var_4, 0000001Ah
  loc_00DF43C5: mov var_38, 00000000h
  loc_00DF43CC: lea edx, var_3C
  loc_00DF43CF: push edx
  loc_00DF43D0: lea eax, var_38
  loc_00DF43D3: push eax
  loc_00DF43D4: push 00FFFFFFh
  loc_00DF43D9: mov ecx, Me
  loc_00DF43DC: mov edx, [ecx]
  loc_00DF43DE: mov eax, Me
  loc_00DF43E1: push eax
  loc_00DF43E2: call [edx+0000090Ch]
  loc_00DF43E8: lea ecx, var_40
  loc_00DF43EB: push ecx
  loc_00DF43EC: mov edx, var_34
  loc_00DF43EF: push edx
  loc_00DF43F0: mov eax, var_3C
  loc_00DF43F3: push eax
  loc_00DF43F4: mov ecx, Me
  loc_00DF43F7: mov edx, [ecx]
  loc_00DF43F9: mov eax, Me
  loc_00DF43FC: push eax
  loc_00DF43FD: call [edx+000008ECh]
  loc_00DF4403: mov ecx, var_40
  loc_00DF4406: push ecx
  loc_00DF4407: mov edx, Me
  loc_00DF440A: mov eax, [edx+000001D0h]
  loc_00DF4410: sub eax, 00000002h
  loc_00DF4413: jo 00DF5F00h
  loc_00DF4419: push eax
  loc_00DF441A: push 00000001h
  loc_00DF441C: push 00000001h
  loc_00DF441E: push 00000001h
  loc_00DF4420: mov ecx, Me
  loc_00DF4423: mov edx, [ecx]
  loc_00DF4425: mov eax, Me
  loc_00DF4428: push eax
  loc_00DF4429: call [edx+000008E4h]
  loc_00DF442F: jmp 00DF46CCh
  loc_00DF4434: mov var_4, 0000001Ch
  loc_00DF443B: mov var_38, 3E6147AEh
  loc_00DF4442: lea ecx, var_3C
  loc_00DF4445: push ecx
  loc_00DF4446: lea edx, var_38
  loc_00DF4449: push edx
  loc_00DF444A: lea eax, var_34
  loc_00DF444D: push eax
  loc_00DF444E: mov ecx, Me
  loc_00DF4451: mov edx, [ecx]
  loc_00DF4453: mov eax, Me
  loc_00DF4456: push eax
  loc_00DF4457: call [edx+0000097Ch]
  loc_00DF445D: push 00000001h
  loc_00DF445F: mov ecx, var_34
  loc_00DF4462: push ecx
  loc_00DF4463: mov edx, var_3C
  loc_00DF4466: push edx
  loc_00DF4467: mov eax, Me
  loc_00DF446A: fild real4 ptr [eax+000001D0h]
  loc_00DF4470: fstp real8 ptr var_94
  loc_00DF4476: fld real8 ptr var_94
  loc_00DF447C: cmp [00E3D000h], 00000000h
  loc_00DF4483: jnz 00DF448Dh
  loc_00DF4485: fdiv st0, real8 ptr [00401338h]
  loc_00DF448B: jmp 00DF449Eh
  loc_00DF448D: push [0040133Ch]
  loc_00DF4493: push [00401338h]
  loc_00DF4499: call 00402854h ; _adj_fdiv_m64
  loc_00DF449E: fnstsw ax
  loc_00DF44A0: test al, 0Dh
  loc_00DF44A2: jnz 00DF5EFBh
  loc_00DF44A8: call [00401208h] ; __vbaFpI4
  loc_00DF44AE: push eax
  loc_00DF44AF: mov ecx, Me
  loc_00DF44B2: mov edx, [ecx+000001D4h]
  loc_00DF44B8: push edx
  loc_00DF44B9: push 00000000h
  loc_00DF44BB: push 00000000h
  loc_00DF44BD: mov eax, Me
  loc_00DF44C0: mov ecx, [eax]
  loc_00DF44C2: mov edx, Me
  loc_00DF44C5: push edx
  loc_00DF44C6: call [ecx+00000908h]
  loc_00DF44CC: mov var_4, 0000001Dh
  loc_00DF44D3: fld real8 ptr [004017C8h]
  loc_00DF44D9: fchs
  loc_00DF44DB: fstp real4 ptr var_38
  loc_00DF44DE: fnstsw ax
  loc_00DF44E0: test al, 0Dh
  loc_00DF44E2: jnz 00DF5EFBh
  loc_00DF44E8: lea eax, var_3C
  loc_00DF44EB: push eax
  loc_00DF44EC: lea ecx, var_38
  loc_00DF44EF: push ecx
  loc_00DF44F0: lea edx, var_34
  loc_00DF44F3: push edx
  loc_00DF44F4: mov eax, Me
  loc_00DF44F7: mov ecx, [eax]
  loc_00DF44F9: mov edx, Me
  loc_00DF44FC: push edx
  loc_00DF44FD: call [ecx+0000097Ch]
  loc_00DF4503: fld real8 ptr [004017B8h]
  loc_00DF4509: fchs
  loc_00DF450B: fstp real4 ptr var_40
  loc_00DF450E: fnstsw ax
  loc_00DF4510: test al, 0Dh
  loc_00DF4512: jnz 00DF5EFBh
  loc_00DF4518: lea eax, var_44
  loc_00DF451B: push eax
  loc_00DF451C: lea ecx, var_40
  loc_00DF451F: push ecx
  loc_00DF4520: lea edx, var_34
  loc_00DF4523: push edx
  loc_00DF4524: mov eax, Me
  loc_00DF4527: mov ecx, [eax]
  loc_00DF4529: mov edx, Me
  loc_00DF452C: push edx
  loc_00DF452D: call [ecx+0000097Ch]
  loc_00DF4533: push 00000001h
  loc_00DF4535: mov eax, var_44
  loc_00DF4538: push eax
  loc_00DF4539: mov ecx, var_3C
  loc_00DF453C: push ecx
  loc_00DF453D: mov edx, Me
  loc_00DF4540: fild real4 ptr [edx+000001D0h]
  loc_00DF4546: fstp real8 ptr var_9C
  loc_00DF454C: fld real8 ptr var_9C
  loc_00DF4552: cmp [00E3D000h], 00000000h
  loc_00DF4559: jnz 00DF4563h
  loc_00DF455B: fdiv st0, real8 ptr [00401338h]
  loc_00DF4561: jmp 00DF4574h
  loc_00DF4563: push [0040133Ch]
  loc_00DF4569: push [00401338h]
  loc_00DF456F: call 00402854h ; _adj_fdiv_m64
  loc_00DF4574: fnstsw ax
  loc_00DF4576: test al, 0Dh
  loc_00DF4578: jnz 00DF5EFBh
  loc_00DF457E: call [00401208h] ; __vbaFpI4
  loc_00DF4584: push eax
  loc_00DF4585: mov eax, Me
  loc_00DF4588: mov ecx, [eax+000001D4h]
  loc_00DF458E: push ecx
  loc_00DF458F: mov edx, Me
  loc_00DF4592: fild real4 ptr [edx+000001D0h]
  loc_00DF4598: fstp real8 ptr var_A4
  loc_00DF459E: fld real8 ptr var_A4
  loc_00DF45A4: cmp [00E3D000h], 00000000h
  loc_00DF45AB: jnz 00DF45B5h
  loc_00DF45AD: fdiv st0, real8 ptr [00401338h]
  loc_00DF45B3: jmp 00DF45C6h
  loc_00DF45B5: push [0040133Ch]
  loc_00DF45BB: push [00401338h]
  loc_00DF45C1: call 00402854h ; _adj_fdiv_m64
  loc_00DF45C6: fnstsw ax
  loc_00DF45C8: test al, 0Dh
  loc_00DF45CA: jnz 00DF5EFBh
  loc_00DF45D0: call [00401208h] ; __vbaFpI4
  loc_00DF45D6: push eax
  loc_00DF45D7: push 00000000h
  loc_00DF45D9: mov eax, Me
  loc_00DF45DC: mov ecx, [eax]
  loc_00DF45DE: mov edx, Me
  loc_00DF45E1: push edx
  loc_00DF45E2: call [ecx+00000908h]
  loc_00DF45E8: mov var_4, 0000001Eh
  loc_00DF45EF: mov eax, Me
  loc_00DF45F2: mov ecx, [eax]
  loc_00DF45F4: mov edx, Me
  loc_00DF45F7: push edx
  loc_00DF45F8: call [ecx+00000920h]
  loc_00DF45FE: mov var_4, 0000001Fh
  loc_00DF4605: mov var_38, 00000000h
  loc_00DF460C: lea eax, var_3C
  loc_00DF460F: push eax
  loc_00DF4610: lea ecx, var_38
  loc_00DF4613: push ecx
  loc_00DF4614: push 00FFFFFFh
  loc_00DF4619: mov edx, Me
  loc_00DF461C: mov eax, [edx]
  loc_00DF461E: mov ecx, Me
  loc_00DF4621: push ecx
  loc_00DF4622: call [eax+0000090Ch]
  loc_00DF4628: mov edx, var_3C
  loc_00DF462B: push edx
  loc_00DF462C: mov eax, Me
  loc_00DF462F: mov ecx, [eax+000001D0h]
  loc_00DF4635: sub ecx, 00000002h
  loc_00DF4638: jo 00DF5F00h
  loc_00DF463E: push ecx
  loc_00DF463F: mov edx, Me
  loc_00DF4642: mov eax, [edx+000001D4h]
  loc_00DF4648: sub eax, 00000002h
  loc_00DF464B: jo 00DF5F00h
  loc_00DF4651: push eax
  loc_00DF4652: push 00000002h
  loc_00DF4654: mov ecx, Me
  loc_00DF4657: mov edx, [ecx+000001D4h]
  loc_00DF465D: sub edx, 00000002h
  loc_00DF4660: jo 00DF5F00h
  loc_00DF4666: push edx
  loc_00DF4667: mov eax, Me
  loc_00DF466A: mov ecx, [eax]
  loc_00DF466C: mov edx, Me
  loc_00DF466F: push edx
  loc_00DF4670: call [ecx+000008E4h]
  loc_00DF4676: mov var_4, 00000020h
  loc_00DF467D: mov var_38, 00000000h
  loc_00DF4684: lea eax, var_3C
  loc_00DF4687: push eax
  loc_00DF4688: lea ecx, var_38
  loc_00DF468B: push ecx
  loc_00DF468C: push 00FFFFFFh
  loc_00DF4691: mov edx, Me
  loc_00DF4694: mov eax, [edx]
  loc_00DF4696: mov ecx, Me
  loc_00DF4699: push ecx
  loc_00DF469A: call [eax+0000090Ch]
  loc_00DF46A0: mov edx, var_3C
  loc_00DF46A3: push edx
  loc_00DF46A4: mov eax, Me
  loc_00DF46A7: mov ecx, [eax+000001D0h]
  loc_00DF46AD: sub ecx, 00000002h
  loc_00DF46B0: jo 00DF5F00h
  loc_00DF46B6: push ecx
  loc_00DF46B7: push 00000001h
  loc_00DF46B9: push 00000001h
  loc_00DF46BB: push 00000001h
  loc_00DF46BD: mov edx, Me
  loc_00DF46C0: mov eax, [edx]
  loc_00DF46C2: mov ecx, Me
  loc_00DF46C5: push ecx
  loc_00DF46C6: call [eax+000008E4h]
  loc_00DF46CC: jmp 00DF51C6h
  loc_00DF46D1: mov var_4, 00000023h
  loc_00DF46D8: mov edx, Me
  loc_00DF46DB: mov eax, [edx+000000CCh]
  loc_00DF46E1: mov var_5C, eax
  loc_00DF46E4: mov ecx, var_5C
  loc_00DF46E7: mov var_A8, ecx
  loc_00DF46ED: cmp var_A8, 00000003h
  loc_00DF46F4: ja 00DF49BFh
  loc_00DF46FA: mov edx, var_A8
  loc_00DF4700: jmp [edx*4+00DF5EBBh]
  loc_00DF4707: jmp 00DF49BFh
  loc_00DF470C: mov var_4, 00000025h
  loc_00DF4713: mov var_38, 3D8F5C29h
  loc_00DF471A: lea eax, var_3C
  loc_00DF471D: push eax
  loc_00DF471E: lea ecx, var_38
  loc_00DF4721: push ecx
  loc_00DF4722: lea edx, var_34
  loc_00DF4725: push edx
  loc_00DF4726: mov eax, Me
  loc_00DF4729: mov ecx, [eax]
  loc_00DF472B: mov edx, Me
  loc_00DF472E: push edx
  loc_00DF472F: call [ecx+0000097Ch]
  loc_00DF4735: push 00000001h
  loc_00DF4737: mov eax, var_34
  loc_00DF473A: push eax
  loc_00DF473B: mov ecx, var_3C
  loc_00DF473E: push ecx
  loc_00DF473F: mov edx, Me
  loc_00DF4742: mov eax, [edx+000001D0h]
  loc_00DF4748: push eax
  loc_00DF4749: mov ecx, Me
  loc_00DF474C: mov edx, [ecx+000001D4h]
  loc_00DF4752: push edx
  loc_00DF4753: push 00000000h
  loc_00DF4755: push 00000000h
  loc_00DF4757: mov eax, Me
  loc_00DF475A: mov ecx, [eax]
  loc_00DF475C: mov edx, Me
  loc_00DF475F: push edx
  loc_00DF4760: call [ecx+00000908h]
  loc_00DF4766: mov var_4, 00000026h
  loc_00DF476D: mov var_38, 3DCCCCCDh
  loc_00DF4774: lea eax, var_3C
  loc_00DF4777: push eax
  loc_00DF4778: lea ecx, var_38
  loc_00DF477B: push ecx
  loc_00DF477C: lea edx, var_34
  loc_00DF477F: push edx
  loc_00DF4780: mov eax, Me
  loc_00DF4783: mov ecx, [eax]
  loc_00DF4785: mov edx, Me
  loc_00DF4788: push edx
  loc_00DF4789: call [ecx+0000097Ch]
  loc_00DF478F: mov var_40, 3DA3D70Ah
  loc_00DF4796: lea eax, var_44
  loc_00DF4799: push eax
  loc_00DF479A: lea ecx, var_40
  loc_00DF479D: push ecx
  loc_00DF479E: lea edx, var_34
  loc_00DF47A1: push edx
  loc_00DF47A2: mov eax, Me
  loc_00DF47A5: mov ecx, [eax]
  loc_00DF47A7: mov edx, Me
  loc_00DF47AA: push edx
  loc_00DF47AB: call [ecx+0000097Ch]
  loc_00DF47B1: push 00000001h
  loc_00DF47B3: mov eax, var_44
  loc_00DF47B6: push eax
  loc_00DF47B7: mov ecx, var_3C
  loc_00DF47BA: push ecx
  loc_00DF47BB: push 00000004h
  loc_00DF47BD: mov edx, Me
  loc_00DF47C0: mov eax, [edx+000001D4h]
  loc_00DF47C6: push eax
  loc_00DF47C7: push 00000000h
  loc_00DF47C9: push 00000000h
  loc_00DF47CB: mov ecx, Me
  loc_00DF47CE: mov edx, [ecx]
  loc_00DF47D0: mov eax, Me
  loc_00DF47D3: push eax
  loc_00DF47D4: call [edx+00000908h]
  loc_00DF47DA: mov var_4, 00000027h
  loc_00DF47E1: mov ecx, Me
  loc_00DF47E4: mov edx, [ecx]
  loc_00DF47E6: mov eax, Me
  loc_00DF47E9: push eax
  loc_00DF47EA: call [edx+00000920h]
  loc_00DF47F0: jmp 00DF49BFh
  loc_00DF47F5: mov var_4, 00000029h
  loc_00DF47FC: mov var_38, 3E6147AEh
  loc_00DF4803: lea ecx, var_3C
  loc_00DF4806: push ecx
  loc_00DF4807: lea edx, var_38
  loc_00DF480A: push edx
  loc_00DF480B: lea eax, var_34
  loc_00DF480E: push eax
  loc_00DF480F: mov ecx, Me
  loc_00DF4812: mov edx, [ecx]
  loc_00DF4814: mov eax, Me
  loc_00DF4817: push eax
  loc_00DF4818: call [edx+0000097Ch]
  loc_00DF481E: push 00000001h
  loc_00DF4820: mov ecx, var_34
  loc_00DF4823: push ecx
  loc_00DF4824: mov edx, var_3C
  loc_00DF4827: push edx
  loc_00DF4828: mov eax, Me
  loc_00DF482B: fild real4 ptr [eax+000001D0h]
  loc_00DF4831: fstp real8 ptr var_B0
  loc_00DF4837: fld real8 ptr var_B0
  loc_00DF483D: cmp [00E3D000h], 00000000h
  loc_00DF4844: jnz 00DF484Eh
  loc_00DF4846: fdiv st0, real8 ptr [00401338h]
  loc_00DF484C: jmp 00DF485Fh
  loc_00DF484E: push [0040133Ch]
  loc_00DF4854: push [00401338h]
  loc_00DF485A: call 00402854h ; _adj_fdiv_m64
  loc_00DF485F: fnstsw ax
  loc_00DF4861: test al, 0Dh
  loc_00DF4863: jnz 00DF5EFBh
  loc_00DF4869: call [00401208h] ; __vbaFpI4
  loc_00DF486F: push eax
  loc_00DF4870: mov ecx, Me
  loc_00DF4873: mov edx, [ecx+000001D4h]
  loc_00DF4879: push edx
  loc_00DF487A: push 00000000h
  loc_00DF487C: push 00000000h
  loc_00DF487E: mov eax, Me
  loc_00DF4881: mov ecx, [eax]
  loc_00DF4883: mov edx, Me
  loc_00DF4886: push edx
  loc_00DF4887: call [ecx+00000908h]
  loc_00DF488D: mov var_4, 0000002Ah
  loc_00DF4894: fld real8 ptr [004017C8h]
  loc_00DF489A: fchs
  loc_00DF489C: fstp real4 ptr var_38
  loc_00DF489F: fnstsw ax
  loc_00DF48A1: test al, 0Dh
  loc_00DF48A3: jnz 00DF5EFBh
  loc_00DF48A9: lea eax, var_3C
  loc_00DF48AC: push eax
  loc_00DF48AD: lea ecx, var_38
  loc_00DF48B0: push ecx
  loc_00DF48B1: lea edx, var_34
  loc_00DF48B4: push edx
  loc_00DF48B5: mov eax, Me
  loc_00DF48B8: mov ecx, [eax]
  loc_00DF48BA: mov edx, Me
  loc_00DF48BD: push edx
  loc_00DF48BE: call [ecx+0000097Ch]
  loc_00DF48C4: fld real8 ptr [004017B8h]
  loc_00DF48CA: fchs
  loc_00DF48CC: fstp real4 ptr var_40
  loc_00DF48CF: fnstsw ax
  loc_00DF48D1: test al, 0Dh
  loc_00DF48D3: jnz 00DF5EFBh
  loc_00DF48D9: lea eax, var_44
  loc_00DF48DC: push eax
  loc_00DF48DD: lea ecx, var_40
  loc_00DF48E0: push ecx
  loc_00DF48E1: lea edx, var_34
  loc_00DF48E4: push edx
  loc_00DF48E5: mov eax, Me
  loc_00DF48E8: mov ecx, [eax]
  loc_00DF48EA: mov edx, Me
  loc_00DF48ED: push edx
  loc_00DF48EE: call [ecx+0000097Ch]
  loc_00DF48F4: push 00000001h
  loc_00DF48F6: mov eax, var_44
  loc_00DF48F9: push eax
  loc_00DF48FA: mov ecx, var_3C
  loc_00DF48FD: push ecx
  loc_00DF48FE: mov edx, Me
  loc_00DF4901: fild real4 ptr [edx+000001D0h]
  loc_00DF4907: fstp real8 ptr var_B8
  loc_00DF490D: fld real8 ptr var_B8
  loc_00DF4913: cmp [00E3D000h], 00000000h
  loc_00DF491A: jnz 00DF4924h
  loc_00DF491C: fdiv st0, real8 ptr [00401338h]
  loc_00DF4922: jmp 00DF4935h
  loc_00DF4924: push [0040133Ch]
  loc_00DF492A: push [00401338h]
  loc_00DF4930: call 00402854h ; _adj_fdiv_m64
  loc_00DF4935: fnstsw ax
  loc_00DF4937: test al, 0Dh
  loc_00DF4939: jnz 00DF5EFBh
  loc_00DF493F: call [00401208h] ; __vbaFpI4
  loc_00DF4945: push eax
  loc_00DF4946: mov eax, Me
  loc_00DF4949: mov ecx, [eax+000001D4h]
  loc_00DF494F: push ecx
  loc_00DF4950: mov edx, Me
  loc_00DF4953: fild real4 ptr [edx+000001D0h]
  loc_00DF4959: fstp real8 ptr var_C0
  loc_00DF495F: fld real8 ptr var_C0
  loc_00DF4965: cmp [00E3D000h], 00000000h
  loc_00DF496C: jnz 00DF4976h
  loc_00DF496E: fdiv st0, real8 ptr [00401338h]
  loc_00DF4974: jmp 00DF4987h
  loc_00DF4976: push [0040133Ch]
  loc_00DF497C: push [00401338h]
  loc_00DF4982: call 00402854h ; _adj_fdiv_m64
  loc_00DF4987: fnstsw ax
  loc_00DF4989: test al, 0Dh
  loc_00DF498B: jnz 00DF5EFBh
  loc_00DF4991: call [00401208h] ; __vbaFpI4
  loc_00DF4997: push eax
  loc_00DF4998: push 00000000h
  loc_00DF499A: mov eax, Me
  loc_00DF499D: mov ecx, [eax]
  loc_00DF499F: mov edx, Me
  loc_00DF49A2: push edx
  loc_00DF49A3: call [ecx+00000908h]
  loc_00DF49A9: mov var_4, 0000002Bh
  loc_00DF49B0: mov eax, Me
  loc_00DF49B3: mov ecx, [eax]
  loc_00DF49B5: mov edx, Me
  loc_00DF49B8: push edx
  loc_00DF49B9: call [ecx+00000920h]
  loc_00DF49BF: mov var_4, 0000002Dh
  loc_00DF49C6: mov var_38, 00000000h
  loc_00DF49CD: lea eax, var_3C
  loc_00DF49D0: push eax
  loc_00DF49D1: lea ecx, var_38
  loc_00DF49D4: push ecx
  loc_00DF49D5: push 0089D8FDh
  loc_00DF49DA: mov edx, Me
  loc_00DF49DD: mov eax, [edx]
  loc_00DF49DF: mov ecx, Me
  loc_00DF49E2: push ecx
  loc_00DF49E3: call [eax+0000090Ch]
  loc_00DF49E9: mov edx, var_3C
  loc_00DF49EC: push edx
  loc_00DF49ED: push 00000002h
  loc_00DF49EF: mov eax, Me
  loc_00DF49F2: mov ecx, [eax+000001D4h]
  loc_00DF49F8: sub ecx, 00000002h
  loc_00DF49FB: jo 00DF5F00h
  loc_00DF4A01: push ecx
  loc_00DF4A02: push 00000002h
  loc_00DF4A04: push 00000001h
  loc_00DF4A06: mov edx, Me
  loc_00DF4A09: mov eax, [edx]
  loc_00DF4A0B: mov ecx, Me
  loc_00DF4A0E: push ecx
  loc_00DF4A0F: call [eax+000008E4h]
  loc_00DF4A15: mov var_4, 0000002Eh
  loc_00DF4A1C: mov var_38, 00000000h
  loc_00DF4A23: lea edx, var_3C
  loc_00DF4A26: push edx
  loc_00DF4A27: lea eax, var_38
  loc_00DF4A2A: push eax
  loc_00DF4A2B: push 00CFF0FFh
  loc_00DF4A30: mov ecx, Me
  loc_00DF4A33: mov edx, [ecx]
  loc_00DF4A35: mov eax, Me
  loc_00DF4A38: push eax
  loc_00DF4A39: call [edx+0000090Ch]
  loc_00DF4A3F: mov ecx, var_3C
  loc_00DF4A42: push ecx
  loc_00DF4A43: push 00000001h
  loc_00DF4A45: mov edx, Me
  loc_00DF4A48: mov eax, [edx+000001D4h]
  loc_00DF4A4E: sub eax, 00000002h
  loc_00DF4A51: jo 00DF5F00h
  loc_00DF4A57: push eax
  loc_00DF4A58: push 00000001h
  loc_00DF4A5A: push 00000001h
  loc_00DF4A5C: mov ecx, Me
  loc_00DF4A5F: mov edx, [ecx]
  loc_00DF4A61: mov eax, Me
  loc_00DF4A64: push eax
  loc_00DF4A65: call [edx+000008E4h]
  loc_00DF4A6B: mov var_4, 0000002Fh
  loc_00DF4A72: mov var_38, 00000000h
  loc_00DF4A79: lea ecx, var_3C
  loc_00DF4A7C: push ecx
  loc_00DF4A7D: lea edx, var_38
  loc_00DF4A80: push edx
  loc_00DF4A81: push 0049BDF9h
  loc_00DF4A86: mov eax, Me
  loc_00DF4A89: mov ecx, [eax]
  loc_00DF4A8B: mov edx, Me
  loc_00DF4A8E: push edx
  loc_00DF4A8F: call [ecx+0000090Ch]
  loc_00DF4A95: mov eax, var_3C
  loc_00DF4A98: push eax
  loc_00DF4A99: mov ecx, Me
  loc_00DF4A9C: mov edx, [ecx+000001D0h]
  loc_00DF4AA2: sub edx, 00000002h
  loc_00DF4AA5: jo 00DF5F00h
  loc_00DF4AAB: push edx
  loc_00DF4AAC: push 00000001h
  loc_00DF4AAE: push 00000001h
  loc_00DF4AB0: push 00000001h
  loc_00DF4AB2: mov eax, Me
  loc_00DF4AB5: mov ecx, [eax]
  loc_00DF4AB7: mov edx, Me
  loc_00DF4ABA: push edx
  loc_00DF4ABB: call [ecx+000008E4h]
  loc_00DF4AC1: mov var_4, 00000030h
  loc_00DF4AC8: mov var_38, 00000000h
  loc_00DF4ACF: lea eax, var_3C
  loc_00DF4AD2: push eax
  loc_00DF4AD3: lea ecx, var_38
  loc_00DF4AD6: push ecx
  loc_00DF4AD7: push 0049BDF9h
  loc_00DF4ADC: mov edx, Me
  loc_00DF4ADF: mov eax, [edx]
  loc_00DF4AE1: mov ecx, Me
  loc_00DF4AE4: push ecx
  loc_00DF4AE5: call [eax+0000090Ch]
  loc_00DF4AEB: mov edx, var_3C
  loc_00DF4AEE: push edx
  loc_00DF4AEF: mov eax, Me
  loc_00DF4AF2: mov ecx, [eax+000001D0h]
  loc_00DF4AF8: sub ecx, 00000002h
  loc_00DF4AFB: jo 00DF5F00h
  loc_00DF4B01: push ecx
  loc_00DF4B02: mov edx, Me
  loc_00DF4B05: mov eax, [edx+000001D4h]
  loc_00DF4B0B: sub eax, 00000002h
  loc_00DF4B0E: jo 00DF5F00h
  loc_00DF4B14: push eax
  loc_00DF4B15: push 00000002h
  loc_00DF4B17: mov ecx, Me
  loc_00DF4B1A: mov edx, [ecx+000001D4h]
  loc_00DF4B20: sub edx, 00000002h
  loc_00DF4B23: jo 00DF5F00h
  loc_00DF4B29: push edx
  loc_00DF4B2A: mov eax, Me
  loc_00DF4B2D: mov ecx, [eax]
  loc_00DF4B2F: mov edx, Me
  loc_00DF4B32: push edx
  loc_00DF4B33: call [ecx+000008E4h]
  loc_00DF4B39: mov var_4, 00000031h
  loc_00DF4B40: mov var_38, 00000000h
  loc_00DF4B47: lea eax, var_3C
  loc_00DF4B4A: push eax
  loc_00DF4B4B: lea ecx, var_38
  loc_00DF4B4E: push ecx
  loc_00DF4B4F: push 007AD2FCh
  loc_00DF4B54: mov edx, Me
  loc_00DF4B57: mov eax, [edx]
  loc_00DF4B59: mov ecx, Me
  loc_00DF4B5C: push ecx
  loc_00DF4B5D: call [eax+0000090Ch]
  loc_00DF4B63: mov edx, var_3C
  loc_00DF4B66: push edx
  loc_00DF4B67: mov eax, Me
  loc_00DF4B6A: mov ecx, [eax+000001D0h]
  loc_00DF4B70: sub ecx, 00000003h
  loc_00DF4B73: jo 00DF5F00h
  loc_00DF4B79: push ecx
  loc_00DF4B7A: push 00000002h
  loc_00DF4B7C: push 00000002h
  loc_00DF4B7E: push 00000002h
  loc_00DF4B80: mov edx, Me
  loc_00DF4B83: mov eax, [edx]
  loc_00DF4B85: mov ecx, Me
  loc_00DF4B88: push ecx
  loc_00DF4B89: call [eax+000008E4h]
  loc_00DF4B8F: mov var_4, 00000032h
  loc_00DF4B96: mov var_38, 00000000h
  loc_00DF4B9D: lea edx, var_3C
  loc_00DF4BA0: push edx
  loc_00DF4BA1: lea eax, var_38
  loc_00DF4BA4: push eax
  loc_00DF4BA5: push 007AD2FCh
  loc_00DF4BAA: mov ecx, Me
  loc_00DF4BAD: mov edx, [ecx]
  loc_00DF4BAF: mov eax, Me
  loc_00DF4BB2: push eax
  loc_00DF4BB3: call [edx+0000090Ch]
  loc_00DF4BB9: mov ecx, var_3C
  loc_00DF4BBC: push ecx
  loc_00DF4BBD: mov edx, Me
  loc_00DF4BC0: mov eax, [edx+000001D0h]
  loc_00DF4BC6: sub eax, 00000003h
  loc_00DF4BC9: jo 00DF5F00h
  loc_00DF4BCF: push eax
  loc_00DF4BD0: mov ecx, Me
  loc_00DF4BD3: mov edx, [ecx+000001D4h]
  loc_00DF4BD9: sub edx, 00000003h
  loc_00DF4BDC: jo 00DF5F00h
  loc_00DF4BE2: push edx
  loc_00DF4BE3: push 00000003h
  loc_00DF4BE5: mov eax, Me
  loc_00DF4BE8: mov ecx, [eax+000001D4h]
  loc_00DF4BEE: sub ecx, 00000003h
  loc_00DF4BF1: jo 00DF5F00h
  loc_00DF4BF7: push ecx
  loc_00DF4BF8: mov edx, Me
  loc_00DF4BFB: mov eax, [edx]
  loc_00DF4BFD: mov ecx, Me
  loc_00DF4C00: push ecx
  loc_00DF4C01: call [eax+000008E4h]
  loc_00DF4C07: mov var_4, 00000033h
  loc_00DF4C0E: mov var_38, 00000000h
  loc_00DF4C15: lea edx, var_3C
  loc_00DF4C18: push edx
  loc_00DF4C19: lea eax, var_38
  loc_00DF4C1C: push eax
  loc_00DF4C1D: push 0030B3F8h
  loc_00DF4C22: mov ecx, Me
  loc_00DF4C25: mov edx, [ecx]
  loc_00DF4C27: mov eax, Me
  loc_00DF4C2A: push eax
  loc_00DF4C2B: call [edx+0000090Ch]
  loc_00DF4C31: mov ecx, var_3C
  loc_00DF4C34: push ecx
  loc_00DF4C35: mov edx, Me
  loc_00DF4C38: mov eax, [edx+000001D0h]
  loc_00DF4C3E: sub eax, 00000003h
  loc_00DF4C41: jo 00DF5F00h
  loc_00DF4C47: push eax
  loc_00DF4C48: mov ecx, Me
  loc_00DF4C4B: mov edx, [ecx+000001D4h]
  loc_00DF4C51: sub edx, 00000002h
  loc_00DF4C54: jo 00DF5F00h
  loc_00DF4C5A: push edx
  loc_00DF4C5B: mov eax, Me
  loc_00DF4C5E: mov ecx, [eax+000001D0h]
  loc_00DF4C64: sub ecx, 00000003h
  loc_00DF4C67: jo 00DF5F00h
  loc_00DF4C6D: push ecx
  loc_00DF4C6E: push 00000002h
  loc_00DF4C70: mov edx, Me
  loc_00DF4C73: mov eax, [edx]
  loc_00DF4C75: mov ecx, Me
  loc_00DF4C78: push ecx
  loc_00DF4C79: call [eax+000008E4h]
  loc_00DF4C7F: mov var_4, 00000034h
  loc_00DF4C86: mov var_38, 00000000h
  loc_00DF4C8D: lea edx, var_3C
  loc_00DF4C90: push edx
  loc_00DF4C91: lea eax, var_38
  loc_00DF4C94: push eax
  loc_00DF4C95: push 000097E5h
  loc_00DF4C9A: mov ecx, Me
  loc_00DF4C9D: mov edx, [ecx]
  loc_00DF4C9F: mov eax, Me
  loc_00DF4CA2: push eax
  loc_00DF4CA3: call [edx+0000090Ch]
  loc_00DF4CA9: mov ecx, var_3C
  loc_00DF4CAC: push ecx
  loc_00DF4CAD: mov edx, Me
  loc_00DF4CB0: mov eax, [edx+000001D0h]
  loc_00DF4CB6: sub eax, 00000002h
  loc_00DF4CB9: jo 00DF5F00h
  loc_00DF4CBF: push eax
  loc_00DF4CC0: mov ecx, Me
  loc_00DF4CC3: mov edx, [ecx+000001D4h]
  loc_00DF4CC9: sub edx, 00000002h
  loc_00DF4CCC: jo 00DF5F00h
  loc_00DF4CD2: push edx
  loc_00DF4CD3: mov eax, Me
  loc_00DF4CD6: mov ecx, [eax+000001D0h]
  loc_00DF4CDC: sub ecx, 00000002h
  loc_00DF4CDF: jo 00DF5F00h
  loc_00DF4CE5: push ecx
  loc_00DF4CE6: push 00000002h
  loc_00DF4CE8: mov edx, Me
  loc_00DF4CEB: mov eax, [edx]
  loc_00DF4CED: mov ecx, Me
  loc_00DF4CF0: push ecx
  loc_00DF4CF1: call [eax+000008E4h]
  loc_00DF4CF7: jmp 00DF51C6h
  loc_00DF4CFC: mov var_4, 00000036h
  loc_00DF4D03: mov edx, Me
  loc_00DF4D06: mov eax, [edx+000000CCh]
  loc_00DF4D0C: mov var_60, eax
  loc_00DF4D0F: mov ecx, var_60
  loc_00DF4D12: mov var_C4, ecx
  loc_00DF4D18: cmp var_C4, 00000003h
  loc_00DF4D1F: ja 00DF51C6h
  loc_00DF4D25: mov edx, var_C4
  loc_00DF4D2B: jmp [edx*4+00DF5ECBh]
  loc_00DF4D32: jmp 00DF51C6h
  loc_00DF4D37: mov var_4, 00000038h
  loc_00DF4D3E: fld real8 ptr [004017D0h]
  loc_00DF4D44: fchs
  loc_00DF4D46: fstp real4 ptr var_38
  loc_00DF4D49: fnstsw ax
  loc_00DF4D4B: test al, 0Dh
  loc_00DF4D4D: jnz 00DF5EFBh
  loc_00DF4D53: lea eax, var_3C
  loc_00DF4D56: push eax
  loc_00DF4D57: lea ecx, var_38
  loc_00DF4D5A: push ecx
  loc_00DF4D5B: lea edx, var_34
  loc_00DF4D5E: push edx
  loc_00DF4D5F: mov eax, Me
  loc_00DF4D62: mov ecx, [eax]
  loc_00DF4D64: mov edx, Me
  loc_00DF4D67: push edx
  loc_00DF4D68: call [ecx+0000097Ch]
  loc_00DF4D6E: mov eax, Me
  loc_00DF4D71: add eax, 000000ACh
  loc_00DF4D76: push eax
  loc_00DF4D77: mov ecx, var_3C
  loc_00DF4D7A: push ecx
  loc_00DF4D7B: mov edx, Me
  loc_00DF4D7E: mov eax, [edx]
  loc_00DF4D80: mov ecx, Me
  loc_00DF4D83: push ecx
  loc_00DF4D84: call [eax+00000974h]
  loc_00DF4D8A: mov var_4, 00000039h
  loc_00DF4D91: mov edx, Me
  loc_00DF4D94: mov eax, [edx]
  loc_00DF4D96: mov ecx, Me
  loc_00DF4D99: push ecx
  loc_00DF4D9A: call [eax+00000920h]
  loc_00DF4DA0: mov var_4, 0000003Ah
  loc_00DF4DA7: fld real8 ptr [004017B0h]
  loc_00DF4DAD: fchs
  loc_00DF4DAF: fstp real4 ptr var_38
  loc_00DF4DB2: fnstsw ax
  loc_00DF4DB4: test al, 0Dh
  loc_00DF4DB6: jnz 00DF5EFBh
  loc_00DF4DBC: lea edx, var_3C
  loc_00DF4DBF: push edx
  loc_00DF4DC0: lea eax, var_38
  loc_00DF4DC3: push eax
  loc_00DF4DC4: lea ecx, var_34
  loc_00DF4DC7: push ecx
  loc_00DF4DC8: mov edx, Me
  loc_00DF4DCB: mov eax, [edx]
  loc_00DF4DCD: mov ecx, Me
  loc_00DF4DD0: push ecx
  loc_00DF4DD1: call [eax+0000097Ch]
  loc_00DF4DD7: mov edx, var_3C
  loc_00DF4DDA: push edx
  loc_00DF4DDB: push 00000001h
  loc_00DF4DDD: mov eax, Me
  loc_00DF4DE0: mov ecx, [eax+000001D4h]
  loc_00DF4DE6: sub ecx, 00000002h
  loc_00DF4DE9: jo 00DF5F00h
  loc_00DF4DEF: push ecx
  loc_00DF4DF0: push 00000001h
  loc_00DF4DF2: push 00000001h
  loc_00DF4DF4: mov edx, Me
  loc_00DF4DF7: mov eax, [edx]
  loc_00DF4DF9: mov ecx, Me
  loc_00DF4DFC: push ecx
  loc_00DF4DFD: call [eax+000008E4h]
  loc_00DF4E03: mov var_4, 0000003Bh
  loc_00DF4E0A: fld real8 ptr [004017E0h]
  loc_00DF4E10: fchs
  loc_00DF4E12: fstp real4 ptr var_38
  loc_00DF4E15: fnstsw ax
  loc_00DF4E17: test al, 0Dh
  loc_00DF4E19: jnz 00DF5EFBh
  loc_00DF4E1F: lea edx, var_3C
  loc_00DF4E22: push edx
  loc_00DF4E23: lea eax, var_38
  loc_00DF4E26: push eax
  loc_00DF4E27: lea ecx, var_34
  loc_00DF4E2A: push ecx
  loc_00DF4E2B: mov edx, Me
  loc_00DF4E2E: mov eax, [edx]
  loc_00DF4E30: mov ecx, Me
  loc_00DF4E33: push ecx
  loc_00DF4E34: call [eax+0000097Ch]
  loc_00DF4E3A: mov edx, var_3C
  loc_00DF4E3D: push edx
  loc_00DF4E3E: push 00000002h
  loc_00DF4E40: mov eax, Me
  loc_00DF4E43: mov ecx, [eax+000001D4h]
  loc_00DF4E49: sub ecx, 00000002h
  loc_00DF4E4C: jo 00DF5F00h
  loc_00DF4E52: push ecx
  loc_00DF4E53: push 00000002h
  loc_00DF4E55: push 00000001h
  loc_00DF4E57: mov edx, Me
  loc_00DF4E5A: mov eax, [edx]
  loc_00DF4E5C: mov ecx, Me
  loc_00DF4E5F: push ecx
  loc_00DF4E60: call [eax+000008E4h]
  loc_00DF4E66: mov var_4, 0000003Ch
  loc_00DF4E6D: mov var_38, 3C23D70Ah
  loc_00DF4E74: lea edx, var_3C
  loc_00DF4E77: push edx
  loc_00DF4E78: lea eax, var_38
  loc_00DF4E7B: push eax
  loc_00DF4E7C: lea ecx, var_34
  loc_00DF4E7F: push ecx
  loc_00DF4E80: mov edx, Me
  loc_00DF4E83: mov eax, [edx]
  loc_00DF4E85: mov ecx, Me
  loc_00DF4E88: push ecx
  loc_00DF4E89: call [eax+0000097Ch]
  loc_00DF4E8F: mov edx, var_3C
  loc_00DF4E92: push edx
  loc_00DF4E93: mov eax, Me
  loc_00DF4E96: mov ecx, [eax+000001D0h]
  loc_00DF4E9C: sub ecx, 00000002h
  loc_00DF4E9F: jo 00DF5F00h
  loc_00DF4EA5: push ecx
  loc_00DF4EA6: mov edx, Me
  loc_00DF4EA9: mov eax, [edx+000001D4h]
  loc_00DF4EAF: sub eax, 00000002h
  loc_00DF4EB2: jo 00DF5F00h
  loc_00DF4EB8: push eax
  loc_00DF4EB9: mov ecx, Me
  loc_00DF4EBC: mov edx, [ecx+000001D0h]
  loc_00DF4EC2: sub edx, 00000002h
  loc_00DF4EC5: jo 00DF5F00h
  loc_00DF4ECB: push edx
  loc_00DF4ECC: push 00000001h
  loc_00DF4ECE: mov eax, Me
  loc_00DF4ED1: mov ecx, [eax]
  loc_00DF4ED3: mov edx, Me
  loc_00DF4ED6: push edx
  loc_00DF4ED7: call [ecx+000008E4h]
  loc_00DF4EDD: mov var_4, 0000003Dh
  loc_00DF4EE4: fld real8 ptr [004017B0h]
  loc_00DF4EEA: fchs
  loc_00DF4EEC: fstp real4 ptr var_38
  loc_00DF4EEF: fnstsw ax
  loc_00DF4EF1: test al, 0Dh
  loc_00DF4EF3: jnz 00DF5EFBh
  loc_00DF4EF9: lea eax, var_3C
  loc_00DF4EFC: push eax
  loc_00DF4EFD: lea ecx, var_38
  loc_00DF4F00: push ecx
  loc_00DF4F01: lea edx, var_34
  loc_00DF4F04: push edx
  loc_00DF4F05: mov eax, Me
  loc_00DF4F08: mov ecx, [eax]
  loc_00DF4F0A: mov edx, Me
  loc_00DF4F0D: push edx
  loc_00DF4F0E: call [ecx+0000097Ch]
  loc_00DF4F14: mov eax, var_3C
  loc_00DF4F17: push eax
  loc_00DF4F18: mov ecx, Me
  loc_00DF4F1B: mov edx, [ecx+000001D0h]
  loc_00DF4F21: sub edx, 00000002h
  loc_00DF4F24: jo 00DF5F00h
  loc_00DF4F2A: push edx
  loc_00DF4F2B: push 00000001h
  loc_00DF4F2D: push 00000001h
  loc_00DF4F2F: push 00000001h
  loc_00DF4F31: mov eax, Me
  loc_00DF4F34: mov ecx, [eax]
  loc_00DF4F36: mov edx, Me
  loc_00DF4F39: push edx
  loc_00DF4F3A: call [ecx+000008E4h]
  loc_00DF4F40: mov var_4, 0000003Eh
  loc_00DF4F47: fld real8 ptr [004017E0h]
  loc_00DF4F4D: fchs
  loc_00DF4F4F: fstp real4 ptr var_38
  loc_00DF4F52: fnstsw ax
  loc_00DF4F54: test al, 0Dh
  loc_00DF4F56: jnz 00DF5EFBh
  loc_00DF4F5C: lea eax, var_3C
  loc_00DF4F5F: push eax
  loc_00DF4F60: lea ecx, var_38
  loc_00DF4F63: push ecx
  loc_00DF4F64: lea edx, var_34
  loc_00DF4F67: push edx
  loc_00DF4F68: mov eax, Me
  loc_00DF4F6B: mov ecx, [eax]
  loc_00DF4F6D: mov edx, Me
  loc_00DF4F70: push edx
  loc_00DF4F71: call [ecx+0000097Ch]
  loc_00DF4F77: mov eax, var_3C
  loc_00DF4F7A: push eax
  loc_00DF4F7B: mov ecx, Me
  loc_00DF4F7E: mov edx, [ecx+000001D0h]
  loc_00DF4F84: sub edx, 00000002h
  loc_00DF4F87: jo 00DF5F00h
  loc_00DF4F8D: push edx
  loc_00DF4F8E: push 00000002h
  loc_00DF4F90: push 00000002h
  loc_00DF4F92: push 00000002h
  loc_00DF4F94: mov eax, Me
  loc_00DF4F97: mov ecx, [eax]
  loc_00DF4F99: mov edx, Me
  loc_00DF4F9C: push edx
  loc_00DF4F9D: call [ecx+000008E4h]
  loc_00DF4FA3: mov var_4, 0000003Fh
  loc_00DF4FAA: mov var_38, 3D23D70Ah
  loc_00DF4FB1: lea eax, var_3C
  loc_00DF4FB4: push eax
  loc_00DF4FB5: lea ecx, var_38
  loc_00DF4FB8: push ecx
  loc_00DF4FB9: lea edx, var_34
  loc_00DF4FBC: push edx
  loc_00DF4FBD: mov eax, Me
  loc_00DF4FC0: mov ecx, [eax]
  loc_00DF4FC2: mov edx, Me
  loc_00DF4FC5: push edx
  loc_00DF4FC6: call [ecx+0000097Ch]
  loc_00DF4FCC: mov eax, var_3C
  loc_00DF4FCF: push eax
  loc_00DF4FD0: mov ecx, Me
  loc_00DF4FD3: mov edx, [ecx+000001D0h]
  loc_00DF4FD9: sub edx, 00000002h
  loc_00DF4FDC: jo 00DF5F00h
  loc_00DF4FE2: push edx
  loc_00DF4FE3: mov eax, Me
  loc_00DF4FE6: mov ecx, [eax+000001D4h]
  loc_00DF4FEC: sub ecx, 00000002h
  loc_00DF4FEF: jo 00DF5F00h
  loc_00DF4FF5: push ecx
  loc_00DF4FF6: push 00000002h
  loc_00DF4FF8: mov edx, Me
  loc_00DF4FFB: mov eax, [edx+000001D4h]
  loc_00DF5001: sub eax, 00000002h
  loc_00DF5004: jo 00DF5F00h
  loc_00DF500A: push eax
  loc_00DF500B: mov ecx, Me
  loc_00DF500E: mov edx, [ecx]
  loc_00DF5010: mov eax, Me
  loc_00DF5013: push eax
  loc_00DF5014: call [edx+000008E4h]
  loc_00DF501A: jmp 00DF51C6h
  loc_00DF501F: mov var_4, 00000041h
  loc_00DF5026: fld real8 ptr [004017A8h]
  loc_00DF502C: fchs
  loc_00DF502E: fstp real4 ptr var_38
  loc_00DF5031: fnstsw ax
  loc_00DF5033: test al, 0Dh
  loc_00DF5035: jnz 00DF5EFBh
  loc_00DF503B: lea ecx, var_3C
  loc_00DF503E: push ecx
  loc_00DF503F: lea edx, var_38
  loc_00DF5042: push edx
  loc_00DF5043: lea eax, var_34
  loc_00DF5046: push eax
  loc_00DF5047: mov ecx, Me
  loc_00DF504A: mov edx, [ecx]
  loc_00DF504C: mov eax, Me
  loc_00DF504F: push eax
  loc_00DF5050: call [edx+0000097Ch]
  loc_00DF5056: mov var_40, 3D4CCCCDh
  loc_00DF505D: lea ecx, var_44
  loc_00DF5060: push ecx
  loc_00DF5061: lea edx, var_40
  loc_00DF5064: push edx
  loc_00DF5065: lea eax, var_34
  loc_00DF5068: push eax
  loc_00DF5069: mov ecx, Me
  loc_00DF506C: mov edx, [ecx]
  loc_00DF506E: mov eax, Me
  loc_00DF5071: push eax
  loc_00DF5072: call [edx+0000097Ch]
  loc_00DF5078: push 00000001h
  loc_00DF507A: mov ecx, var_44
  loc_00DF507D: push ecx
  loc_00DF507E: mov edx, var_3C
  loc_00DF5081: push edx
  loc_00DF5082: mov eax, Me
  loc_00DF5085: mov ecx, [eax+000001D0h]
  loc_00DF508B: sub ecx, 00000006h
  loc_00DF508E: jo 00DF5F00h
  loc_00DF5094: push ecx
  loc_00DF5095: mov edx, Me
  loc_00DF5098: mov eax, [edx+000001D4h]
  loc_00DF509E: push eax
  loc_00DF509F: push 00000000h
  loc_00DF50A1: push 00000000h
  loc_00DF50A3: mov ecx, Me
  loc_00DF50A6: mov edx, [ecx]
  loc_00DF50A8: mov eax, Me
  loc_00DF50AB: push eax
  loc_00DF50AC: call [edx+00000908h]
  loc_00DF50B2: mov var_4, 00000042h
  loc_00DF50B9: mov var_38, 3DA3D70Ah
  loc_00DF50C0: lea ecx, var_3C
  loc_00DF50C3: push ecx
  loc_00DF50C4: lea edx, var_38
  loc_00DF50C7: push edx
  loc_00DF50C8: lea eax, var_34
  loc_00DF50CB: push eax
  loc_00DF50CC: mov ecx, Me
  loc_00DF50CF: mov edx, [ecx]
  loc_00DF50D1: mov eax, Me
  loc_00DF50D4: push eax
  loc_00DF50D5: call [edx+0000097Ch]
  loc_00DF50DB: mov var_40, 00000000h
  loc_00DF50E2: lea ecx, var_44
  loc_00DF50E5: push ecx
  loc_00DF50E6: lea edx, var_40
  loc_00DF50E9: push edx
  loc_00DF50EA: push 00FFFFFFh
  loc_00DF50EF: mov eax, Me
  loc_00DF50F2: mov ecx, [eax]
  loc_00DF50F4: mov edx, Me
  loc_00DF50F7: push edx
  loc_00DF50F8: call [ecx+0000090Ch]
  loc_00DF50FE: push 00000001h
  loc_00DF5100: mov eax, var_44
  loc_00DF5103: push eax
  loc_00DF5104: mov ecx, var_3C
  loc_00DF5107: push ecx
  loc_00DF5108: mov edx, Me
  loc_00DF510B: mov eax, [edx+000001D0h]
  loc_00DF5111: sub eax, 00000001h
  loc_00DF5114: jo 00DF5F00h
  loc_00DF511A: push eax
  loc_00DF511B: mov ecx, Me
  loc_00DF511E: mov edx, [ecx+000001D4h]
  loc_00DF5124: push edx
  loc_00DF5125: mov eax, Me
  loc_00DF5128: mov ecx, [eax+000001D0h]
  loc_00DF512E: sub ecx, 00000006h
  loc_00DF5131: jo 00DF5F00h
  loc_00DF5137: push ecx
  loc_00DF5138: push 00000000h
  loc_00DF513A: mov edx, Me
  loc_00DF513D: mov eax, [edx]
  loc_00DF513F: mov ecx, Me
  loc_00DF5142: push ecx
  loc_00DF5143: call [eax+00000908h]
  loc_00DF5149: mov var_4, 00000043h
  loc_00DF5150: mov edx, Me
  loc_00DF5153: mov eax, [edx]
  loc_00DF5155: mov ecx, Me
  loc_00DF5158: push ecx
  loc_00DF5159: call [eax+00000920h]
  loc_00DF515F: mov var_4, 00000044h
  loc_00DF5166: mov var_38, 00000000h
  loc_00DF516D: lea edx, var_3C
  loc_00DF5170: push edx
  loc_00DF5171: lea eax, var_38
  loc_00DF5174: push eax
  loc_00DF5175: push 00FFFFFFh
  loc_00DF517A: mov ecx, Me
  loc_00DF517D: mov edx, [ecx]
  loc_00DF517F: mov eax, Me
  loc_00DF5182: push eax
  loc_00DF5183: call [edx+0000090Ch]
  loc_00DF5189: mov ecx, var_3C
  loc_00DF518C: push ecx
  loc_00DF518D: mov edx, Me
  loc_00DF5190: mov eax, [edx+000001D0h]
  loc_00DF5196: sub eax, 00000002h
  loc_00DF5199: jo 00DF5F00h
  loc_00DF519F: push eax
  loc_00DF51A0: mov ecx, Me
  loc_00DF51A3: mov edx, [ecx+000001D4h]
  loc_00DF51A9: sub edx, 00000002h
  loc_00DF51AC: jo 00DF5F00h
  loc_00DF51B2: push edx
  loc_00DF51B3: push 00000001h
  loc_00DF51B5: push 00000001h
  loc_00DF51B7: mov eax, Me
  loc_00DF51BA: mov ecx, [eax]
  loc_00DF51BC: mov edx, Me
  loc_00DF51BF: push edx
  loc_00DF51C0: call [ecx+000008F0h]
  loc_00DF51C6: mov var_4, 00000047h
  loc_00DF51CD: mov eax, Me
  loc_00DF51D0: movsx ecx, [eax+00000070h]
  loc_00DF51D4: test ecx, ecx
  loc_00DF51D6: jz 00DF5BFEh
  loc_00DF51DC: mov var_4, 00000048h
  loc_00DF51E3: mov edx, Me
  loc_00DF51E6: movsx eax, [edx+00000050h]
  loc_00DF51EA: test eax, eax
  loc_00DF51EC: jnz 00DF51FDh
  loc_00DF51EE: mov ecx, Me
  loc_00DF51F1: movsx edx, [ecx+00000058h]
  loc_00DF51F5: test edx, edx
  loc_00DF51F7: jz 00DF5BFEh
  loc_00DF51FD: cmp arg_C, 00000002h
  loc_00DF5201: jz 00DF5BFEh
  loc_00DF5207: cmp arg_C, 00000001h
  loc_00DF520B: jz 00DF5BFEh
  loc_00DF5211: mov var_4, 00000049h
  loc_00DF5218: mov eax, Me
  loc_00DF521B: mov ecx, [eax+000000CCh]
  loc_00DF5221: mov var_64, ecx
  loc_00DF5224: mov edx, var_64
  loc_00DF5227: mov var_C8, edx
  loc_00DF522D: cmp var_C8, 00000003h
  loc_00DF5234: ja 00DF5BFEh
  loc_00DF523A: mov eax, var_C8
  loc_00DF5240: jmp [eax*4+00DF5EDBh]
  loc_00DF5247: jmp 00DF5BFEh
  loc_00DF524C: mov var_4, 0000004Bh
  loc_00DF5253: mov var_38, 00000000h
  loc_00DF525A: lea ecx, var_3C
  loc_00DF525D: push ecx
  loc_00DF525E: lea edx, var_38
  loc_00DF5261: push edx
  loc_00DF5262: push 00F6D4BCh
  loc_00DF5267: mov eax, Me
  loc_00DF526A: mov ecx, [eax]
  loc_00DF526C: mov edx, Me
  loc_00DF526F: push edx
  loc_00DF5270: call [ecx+0000090Ch]
  loc_00DF5276: mov eax, var_3C
  loc_00DF5279: push eax
  loc_00DF527A: push 00000002h
  loc_00DF527C: mov ecx, Me
  loc_00DF527F: mov edx, [ecx+000001D4h]
  loc_00DF5285: sub edx, 00000002h
  loc_00DF5288: jo 00DF5F00h
  loc_00DF528E: push edx
  loc_00DF528F: push 00000002h
  loc_00DF5291: push 00000001h
  loc_00DF5293: mov eax, Me
  loc_00DF5296: mov ecx, [eax]
  loc_00DF5298: mov edx, Me
  loc_00DF529B: push edx
  loc_00DF529C: call [ecx+000008E4h]
  loc_00DF52A2: mov var_4, 0000004Ch
  loc_00DF52A9: mov var_38, 00000000h
  loc_00DF52B0: lea eax, var_3C
  loc_00DF52B3: push eax
  loc_00DF52B4: lea ecx, var_38
  loc_00DF52B7: push ecx
  loc_00DF52B8: push 00FFE7CEh
  loc_00DF52BD: mov edx, Me
  loc_00DF52C0: mov eax, [edx]
  loc_00DF52C2: mov ecx, Me
  loc_00DF52C5: push ecx
  loc_00DF52C6: call [eax+0000090Ch]
  loc_00DF52CC: mov edx, var_3C
  loc_00DF52CF: push edx
  loc_00DF52D0: push 00000001h
  loc_00DF52D2: mov eax, Me
  loc_00DF52D5: mov ecx, [eax+000001D4h]
  loc_00DF52DB: sub ecx, 00000002h
  loc_00DF52DE: jo 00DF5F00h
  loc_00DF52E4: push ecx
  loc_00DF52E5: push 00000001h
  loc_00DF52E7: push 00000001h
  loc_00DF52E9: mov edx, Me
  loc_00DF52EC: mov eax, [edx]
  loc_00DF52EE: mov ecx, Me
  loc_00DF52F1: push ecx
  loc_00DF52F2: call [eax+000008E4h]
  loc_00DF52F8: mov var_4, 0000004Dh
  loc_00DF52FF: mov var_38, 00000000h
  loc_00DF5306: lea edx, var_3C
  loc_00DF5309: push edx
  loc_00DF530A: lea eax, var_38
  loc_00DF530D: push eax
  loc_00DF530E: push 00E6AF8Eh
  loc_00DF5313: mov ecx, Me
  loc_00DF5316: mov edx, [ecx]
  loc_00DF5318: mov eax, Me
  loc_00DF531B: push eax
  loc_00DF531C: call [edx+0000090Ch]
  loc_00DF5322: mov ecx, var_3C
  loc_00DF5325: push ecx
  loc_00DF5326: mov edx, Me
  loc_00DF5329: mov eax, [edx+000001D0h]
  loc_00DF532F: sub eax, 00000002h
  loc_00DF5332: jo 00DF5F00h
  loc_00DF5338: push eax
  loc_00DF5339: push 00000001h
  loc_00DF533B: push 00000001h
  loc_00DF533D: push 00000001h
  loc_00DF533F: mov ecx, Me
  loc_00DF5342: mov edx, [ecx]
  loc_00DF5344: mov eax, Me
  loc_00DF5347: push eax
  loc_00DF5348: call [edx+000008E4h]
  loc_00DF534E: mov var_4, 0000004Eh
  loc_00DF5355: mov var_38, 00000000h
  loc_00DF535C: lea ecx, var_3C
  loc_00DF535F: push ecx
  loc_00DF5360: lea edx, var_38
  loc_00DF5363: push edx
  loc_00DF5364: push 00E6AF8Eh
  loc_00DF5369: mov eax, Me
  loc_00DF536C: mov ecx, [eax]
  loc_00DF536E: mov edx, Me
  loc_00DF5371: push edx
  loc_00DF5372: call [ecx+0000090Ch]
  loc_00DF5378: mov eax, var_3C
  loc_00DF537B: push eax
  loc_00DF537C: mov ecx, Me
  loc_00DF537F: mov edx, [ecx+000001D0h]
  loc_00DF5385: sub edx, 00000002h
  loc_00DF5388: jo 00DF5F00h
  loc_00DF538E: push edx
  loc_00DF538F: mov eax, Me
  loc_00DF5392: mov ecx, [eax+000001D4h]
  loc_00DF5398: sub ecx, 00000002h
  loc_00DF539B: jo 00DF5F00h
  loc_00DF53A1: push ecx
  loc_00DF53A2: push 00000002h
  loc_00DF53A4: mov edx, Me
  loc_00DF53A7: mov eax, [edx+000001D4h]
  loc_00DF53AD: sub eax, 00000002h
  loc_00DF53B0: jo 00DF5F00h
  loc_00DF53B6: push eax
  loc_00DF53B7: mov ecx, Me
  loc_00DF53BA: mov edx, [ecx]
  loc_00DF53BC: mov eax, Me
  loc_00DF53BF: push eax
  loc_00DF53C0: call [edx+000008E4h]
  loc_00DF53C6: mov var_4, 0000004Fh
  loc_00DF53CD: mov var_38, 00000000h
  loc_00DF53D4: lea ecx, var_3C
  loc_00DF53D7: push ecx
  loc_00DF53D8: lea edx, var_38
  loc_00DF53DB: push edx
  loc_00DF53DC: push 00F4D1B8h
  loc_00DF53E1: mov eax, Me
  loc_00DF53E4: mov ecx, [eax]
  loc_00DF53E6: mov edx, Me
  loc_00DF53E9: push edx
  loc_00DF53EA: call [ecx+0000090Ch]
  loc_00DF53F0: mov eax, var_3C
  loc_00DF53F3: push eax
  loc_00DF53F4: mov ecx, Me
  loc_00DF53F7: mov edx, [ecx+000001D0h]
  loc_00DF53FD: sub edx, 00000003h
  loc_00DF5400: jo 00DF5F00h
  loc_00DF5406: push edx
  loc_00DF5407: push 00000002h
  loc_00DF5409: push 00000002h
  loc_00DF540B: push 00000002h
  loc_00DF540D: mov eax, Me
  loc_00DF5410: mov ecx, [eax]
  loc_00DF5412: mov edx, Me
  loc_00DF5415: push edx
  loc_00DF5416: call [ecx+000008E4h]
  loc_00DF541C: mov var_4, 00000050h
  loc_00DF5423: mov var_38, 00000000h
  loc_00DF542A: lea eax, var_3C
  loc_00DF542D: push eax
  loc_00DF542E: lea ecx, var_38
  loc_00DF5431: push ecx
  loc_00DF5432: push 00F4D1B8h
  loc_00DF5437: mov edx, Me
  loc_00DF543A: mov eax, [edx]
  loc_00DF543C: mov ecx, Me
  loc_00DF543F: push ecx
  loc_00DF5440: call [eax+0000090Ch]
  loc_00DF5446: mov edx, var_3C
  loc_00DF5449: push edx
  loc_00DF544A: mov eax, Me
  loc_00DF544D: mov ecx, [eax+000001D0h]
  loc_00DF5453: sub ecx, 00000003h
  loc_00DF5456: jo 00DF5F00h
  loc_00DF545C: push ecx
  loc_00DF545D: mov edx, Me
  loc_00DF5460: mov eax, [edx+000001D4h]
  loc_00DF5466: sub eax, 00000003h
  loc_00DF5469: jo 00DF5F00h
  loc_00DF546F: push eax
  loc_00DF5470: push 00000003h
  loc_00DF5472: mov ecx, Me
  loc_00DF5475: mov edx, [ecx+000001D4h]
  loc_00DF547B: sub edx, 00000003h
  loc_00DF547E: jo 00DF5F00h
  loc_00DF5484: push edx
  loc_00DF5485: mov eax, Me
  loc_00DF5488: mov ecx, [eax]
  loc_00DF548A: mov edx, Me
  loc_00DF548D: push edx
  loc_00DF548E: call [ecx+000008E4h]
  loc_00DF5494: mov var_4, 00000051h
  loc_00DF549B: mov var_38, 00000000h
  loc_00DF54A2: lea eax, var_3C
  loc_00DF54A5: push eax
  loc_00DF54A6: lea ecx, var_38
  loc_00DF54A9: push ecx
  loc_00DF54AA: push 00E4AD89h
  loc_00DF54AF: mov edx, Me
  loc_00DF54B2: mov eax, [edx]
  loc_00DF54B4: mov ecx, Me
  loc_00DF54B7: push ecx
  loc_00DF54B8: call [eax+0000090Ch]
  loc_00DF54BE: mov edx, var_3C
  loc_00DF54C1: push edx
  loc_00DF54C2: mov eax, Me
  loc_00DF54C5: mov ecx, [eax+000001D0h]
  loc_00DF54CB: sub ecx, 00000003h
  loc_00DF54CE: jo 00DF5F00h
  loc_00DF54D4: push ecx
  loc_00DF54D5: mov edx, Me
  loc_00DF54D8: mov eax, [edx+000001D4h]
  loc_00DF54DE: sub eax, 00000002h
  loc_00DF54E1: jo 00DF5F00h
  loc_00DF54E7: push eax
  loc_00DF54E8: mov ecx, Me
  loc_00DF54EB: mov edx, [ecx+000001D0h]
  loc_00DF54F1: sub edx, 00000003h
  loc_00DF54F4: jo 00DF5F00h
  loc_00DF54FA: push edx
  loc_00DF54FB: push 00000002h
  loc_00DF54FD: mov eax, Me
  loc_00DF5500: mov ecx, [eax]
  loc_00DF5502: mov edx, Me
  loc_00DF5505: push edx
  loc_00DF5506: call [ecx+000008E4h]
  loc_00DF550C: mov var_4, 00000052h
  loc_00DF5513: mov var_38, 00000000h
  loc_00DF551A: lea eax, var_3C
  loc_00DF551D: push eax
  loc_00DF551E: lea ecx, var_38
  loc_00DF5521: push ecx
  loc_00DF5522: push 00EE8269h
  loc_00DF5527: mov edx, Me
  loc_00DF552A: mov eax, [edx]
  loc_00DF552C: mov ecx, Me
  loc_00DF552F: push ecx
  loc_00DF5530: call [eax+0000090Ch]
  loc_00DF5536: mov edx, var_3C
  loc_00DF5539: push edx
  loc_00DF553A: mov eax, Me
  loc_00DF553D: mov ecx, [eax+000001D0h]
  loc_00DF5543: sub ecx, 00000002h
  loc_00DF5546: jo 00DF5F00h
  loc_00DF554C: push ecx
  loc_00DF554D: mov edx, Me
  loc_00DF5550: mov eax, [edx+000001D4h]
  loc_00DF5556: sub eax, 00000002h
  loc_00DF5559: jo 00DF5F00h
  loc_00DF555F: push eax
  loc_00DF5560: mov ecx, Me
  loc_00DF5563: mov edx, [ecx+000001D0h]
  loc_00DF5569: sub edx, 00000002h
  loc_00DF556C: jo 00DF5F00h
  loc_00DF5572: push edx
  loc_00DF5573: push 00000002h
  loc_00DF5575: mov eax, Me
  loc_00DF5578: mov ecx, [eax]
  loc_00DF557A: mov edx, Me
  loc_00DF557D: push edx
  loc_00DF557E: call [ecx+000008E4h]
  loc_00DF5584: jmp 00DF5BFEh
  loc_00DF5589: mov var_4, 00000054h
  loc_00DF5590: mov var_38, 00000000h
  loc_00DF5597: lea eax, var_3C
  loc_00DF559A: push eax
  loc_00DF559B: lea ecx, var_38
  loc_00DF559E: push ecx
  loc_00DF559F: push 008FD1C2h ; "RHpé‹s"
  loc_00DF55A4: mov edx, Me
  loc_00DF55A7: mov eax, [edx]
  loc_00DF55A9: mov ecx, Me
  loc_00DF55AC: push ecx
  loc_00DF55AD: call [eax+0000090Ch]
  loc_00DF55B3: mov edx, var_3C
  loc_00DF55B6: push edx
  loc_00DF55B7: push 00000002h
  loc_00DF55B9: mov eax, Me
  loc_00DF55BC: mov ecx, [eax+000001D4h]
  loc_00DF55C2: sub ecx, 00000002h
  loc_00DF55C5: jo 00DF5F00h
  loc_00DF55CB: push ecx
  loc_00DF55CC: push 00000002h
  loc_00DF55CE: push 00000001h
  loc_00DF55D0: mov edx, Me
  loc_00DF55D3: mov eax, [edx]
  loc_00DF55D5: mov ecx, Me
  loc_00DF55D8: push ecx
  loc_00DF55D9: call [eax+000008E4h]
  loc_00DF55DF: mov var_4, 00000055h
  loc_00DF55E6: mov var_38, 00000000h
  loc_00DF55ED: lea edx, var_3C
  loc_00DF55F0: push edx
  loc_00DF55F1: lea eax, var_38
  loc_00DF55F4: push eax
  loc_00DF55F5: push 0080CBB1h
  loc_00DF55FA: mov ecx, Me
  loc_00DF55FD: mov edx, [ecx]
  loc_00DF55FF: mov eax, Me
  loc_00DF5602: push eax
  loc_00DF5603: call [edx+0000090Ch]
  loc_00DF5609: mov ecx, var_3C
  loc_00DF560C: push ecx
  loc_00DF560D: push 00000001h
  loc_00DF560F: mov edx, Me
  loc_00DF5612: mov eax, [edx+000001D4h]
  loc_00DF5618: sub eax, 00000002h
  loc_00DF561B: jo 00DF5F00h
  loc_00DF5621: push eax
  loc_00DF5622: push 00000001h
  loc_00DF5624: push 00000001h
  loc_00DF5626: mov ecx, Me
  loc_00DF5629: mov edx, [ecx]
  loc_00DF562B: mov eax, Me
  loc_00DF562E: push eax
  loc_00DF562F: call [edx+000008E4h]
  loc_00DF5635: mov var_4, 00000056h
  loc_00DF563C: mov var_38, 00000000h
  loc_00DF5643: lea ecx, var_3C
  loc_00DF5646: push ecx
  loc_00DF5647: lea edx, var_38
  loc_00DF564A: push edx
  loc_00DF564B: push 0068C8A0h
  loc_00DF5650: mov eax, Me
  loc_00DF5653: mov ecx, [eax]
  loc_00DF5655: mov edx, Me
  loc_00DF5658: push edx
  loc_00DF5659: call [ecx+0000090Ch]
  loc_00DF565F: mov eax, var_3C
  loc_00DF5662: push eax
  loc_00DF5663: mov ecx, Me
  loc_00DF5666: mov edx, [ecx+000001D0h]
  loc_00DF566C: sub edx, 00000002h
  loc_00DF566F: jo 00DF5F00h
  loc_00DF5675: push edx
  loc_00DF5676: push 00000001h
  loc_00DF5678: push 00000001h
  loc_00DF567A: push 00000001h
  loc_00DF567C: mov eax, Me
  loc_00DF567F: mov ecx, [eax]
  loc_00DF5681: mov edx, Me
  loc_00DF5684: push edx
  loc_00DF5685: call [ecx+000008E4h]
  loc_00DF568B: mov var_4, 00000057h
  loc_00DF5692: mov var_38, 00000000h
  loc_00DF5699: lea eax, var_3C
  loc_00DF569C: push eax
  loc_00DF569D: lea ecx, var_38
  loc_00DF56A0: push ecx
  loc_00DF56A1: push 0068C8A0h
  loc_00DF56A6: mov edx, Me
  loc_00DF56A9: mov eax, [edx]
  loc_00DF56AB: mov ecx, Me
  loc_00DF56AE: push ecx
  loc_00DF56AF: call [eax+0000090Ch]
  loc_00DF56B5: mov edx, var_3C
  loc_00DF56B8: push edx
  loc_00DF56B9: mov eax, Me
  loc_00DF56BC: mov ecx, [eax+000001D0h]
  loc_00DF56C2: sub ecx, 00000002h
  loc_00DF56C5: jo 00DF5F00h
  loc_00DF56CB: push ecx
  loc_00DF56CC: mov edx, Me
  loc_00DF56CF: mov eax, [edx+000001D4h]
  loc_00DF56D5: sub eax, 00000002h
  loc_00DF56D8: jo 00DF5F00h
  loc_00DF56DE: push eax
  loc_00DF56DF: push 00000002h
  loc_00DF56E1: mov ecx, Me
  loc_00DF56E4: mov edx, [ecx+000001D4h]
  loc_00DF56EA: sub edx, 00000002h
  loc_00DF56ED: jo 00DF5F00h
  loc_00DF56F3: push edx
  loc_00DF56F4: mov eax, Me
  loc_00DF56F7: mov ecx, [eax]
  loc_00DF56F9: mov edx, Me
  loc_00DF56FC: push edx
  loc_00DF56FD: call [ecx+000008E4h]
  loc_00DF5703: mov var_4, 00000058h
  loc_00DF570A: mov var_38, 00000000h
  loc_00DF5711: lea eax, var_3C
  loc_00DF5714: push eax
  loc_00DF5715: lea ecx, var_38
  loc_00DF5718: push ecx
  loc_00DF5719: push 0068C8A0h
  loc_00DF571E: mov edx, Me
  loc_00DF5721: mov eax, [edx]
  loc_00DF5723: mov ecx, Me
  loc_00DF5726: push ecx
  loc_00DF5727: call [eax+0000090Ch]
  loc_00DF572D: mov edx, var_3C
  loc_00DF5730: push edx
  loc_00DF5731: mov eax, Me
  loc_00DF5734: mov ecx, [eax+000001D0h]
  loc_00DF573A: sub ecx, 00000003h
  loc_00DF573D: jo 00DF5F00h
  loc_00DF5743: push ecx
  loc_00DF5744: push 00000002h
  loc_00DF5746: push 00000002h
  loc_00DF5748: push 00000002h
  loc_00DF574A: mov edx, Me
  loc_00DF574D: mov eax, [edx]
  loc_00DF574F: mov ecx, Me
  loc_00DF5752: push ecx
  loc_00DF5753: call [eax+000008E4h]
  loc_00DF5759: mov var_4, 00000059h
  loc_00DF5760: mov var_38, 00000000h
  loc_00DF5767: lea edx, var_3C
  loc_00DF576A: push edx
  loc_00DF576B: lea eax, var_38
  loc_00DF576E: push eax
  loc_00DF576F: push 0068C8A0h
  loc_00DF5774: mov ecx, Me
  loc_00DF5777: mov edx, [ecx]
  loc_00DF5779: mov eax, Me
  loc_00DF577C: push eax
  loc_00DF577D: call [edx+0000090Ch]
  loc_00DF5783: mov ecx, var_3C
  loc_00DF5786: push ecx
  loc_00DF5787: mov edx, Me
  loc_00DF578A: mov eax, [edx+000001D0h]
  loc_00DF5790: sub eax, 00000003h
  loc_00DF5793: jo 00DF5F00h
  loc_00DF5799: push eax
  loc_00DF579A: mov ecx, Me
  loc_00DF579D: mov edx, [ecx+000001D4h]
  loc_00DF57A3: sub edx, 00000003h
  loc_00DF57A6: jo 00DF5F00h
  loc_00DF57AC: push edx
  loc_00DF57AD: push 00000003h
  loc_00DF57AF: mov eax, Me
  loc_00DF57B2: mov ecx, [eax+000001D4h]
  loc_00DF57B8: sub ecx, 00000003h
  loc_00DF57BB: jo 00DF5F00h
  loc_00DF57C1: push ecx
  loc_00DF57C2: mov edx, Me
  loc_00DF57C5: mov eax, [edx]
  loc_00DF57C7: mov ecx, Me
  loc_00DF57CA: push ecx
  loc_00DF57CB: call [eax+000008E4h]
  loc_00DF57D1: mov var_4, 0000005Ah
  loc_00DF57D8: mov var_38, 00000000h
  loc_00DF57DF: lea edx, var_3C
  loc_00DF57E2: push edx
  loc_00DF57E3: lea eax, var_38
  loc_00DF57E6: push eax
  loc_00DF57E7: push 0068C8A0h
  loc_00DF57EC: mov ecx, Me
  loc_00DF57EF: mov edx, [ecx]
  loc_00DF57F1: mov eax, Me
  loc_00DF57F4: push eax
  loc_00DF57F5: call [edx+0000090Ch]
  loc_00DF57FB: mov ecx, var_3C
  loc_00DF57FE: push ecx
  loc_00DF57FF: mov edx, Me
  loc_00DF5802: mov eax, [edx+000001D0h]
  loc_00DF5808: sub eax, 00000003h
  loc_00DF580B: jo 00DF5F00h
  loc_00DF5811: push eax
  loc_00DF5812: mov ecx, Me
  loc_00DF5815: mov edx, [ecx+000001D4h]
  loc_00DF581B: sub edx, 00000002h
  loc_00DF581E: jo 00DF5F00h
  loc_00DF5824: push edx
  loc_00DF5825: mov eax, Me
  loc_00DF5828: mov ecx, [eax+000001D0h]
  loc_00DF582E: sub ecx, 00000003h
  loc_00DF5831: jo 00DF5F00h
  loc_00DF5837: push ecx
  loc_00DF5838: push 00000002h
  loc_00DF583A: mov edx, Me
  loc_00DF583D: mov eax, [edx]
  loc_00DF583F: mov ecx, Me
  loc_00DF5842: push ecx
  loc_00DF5843: call [eax+000008E4h]
  loc_00DF5849: mov var_4, 0000005Bh
  loc_00DF5850: mov var_38, 00000000h
  loc_00DF5857: lea edx, var_3C
  loc_00DF585A: push edx
  loc_00DF585B: lea eax, var_38
  loc_00DF585E: push eax
  loc_00DF585F: push 0066A7A8h
  loc_00DF5864: mov ecx, Me
  loc_00DF5867: mov edx, [ecx]
  loc_00DF5869: mov eax, Me
  loc_00DF586C: push eax
  loc_00DF586D: call [edx+0000090Ch]
  loc_00DF5873: mov ecx, var_3C
  loc_00DF5876: push ecx
  loc_00DF5877: mov edx, Me
  loc_00DF587A: mov eax, [edx+000001D0h]
  loc_00DF5880: sub eax, 00000002h
  loc_00DF5883: jo 00DF5F00h
  loc_00DF5889: push eax
  loc_00DF588A: mov ecx, Me
  loc_00DF588D: mov edx, [ecx+000001D4h]
  loc_00DF5893: sub edx, 00000002h
  loc_00DF5896: jo 00DF5F00h
  loc_00DF589C: push edx
  loc_00DF589D: mov eax, Me
  loc_00DF58A0: mov ecx, [eax+000001D0h]
  loc_00DF58A6: sub ecx, 00000002h
  loc_00DF58A9: jo 00DF5F00h
  loc_00DF58AF: push ecx
  loc_00DF58B0: push 00000002h
  loc_00DF58B2: mov edx, Me
  loc_00DF58B5: mov eax, [edx]
  loc_00DF58B7: mov ecx, Me
  loc_00DF58BA: push ecx
  loc_00DF58BB: call [eax+000008E4h]
  loc_00DF58C1: jmp 00DF5BFEh
  loc_00DF58C6: mov var_4, 0000005Dh
  loc_00DF58CD: mov var_38, 00000000h
  loc_00DF58D4: lea edx, var_3C
  loc_00DF58D7: push edx
  loc_00DF58D8: lea eax, var_38
  loc_00DF58DB: push eax
  loc_00DF58DC: push 00F6D4BCh
  loc_00DF58E1: mov ecx, Me
  loc_00DF58E4: mov edx, [ecx]
  loc_00DF58E6: mov eax, Me
  loc_00DF58E9: push eax
  loc_00DF58EA: call [edx+0000090Ch]
  loc_00DF58F0: mov ecx, var_3C
  loc_00DF58F3: push ecx
  loc_00DF58F4: push 00000002h
  loc_00DF58F6: mov edx, Me
  loc_00DF58F9: mov eax, [edx+000001D4h]
  loc_00DF58FF: sub eax, 00000002h
  loc_00DF5902: jo 00DF5F00h
  loc_00DF5908: push eax
  loc_00DF5909: push 00000002h
  loc_00DF590B: push 00000001h
  loc_00DF590D: mov ecx, Me
  loc_00DF5910: mov edx, [ecx]
  loc_00DF5912: mov eax, Me
  loc_00DF5915: push eax
  loc_00DF5916: call [edx+000008E4h]
  loc_00DF591C: mov var_4, 0000005Eh
  loc_00DF5923: mov var_38, 00000000h
  loc_00DF592A: lea ecx, var_3C
  loc_00DF592D: push ecx
  loc_00DF592E: lea edx, var_38
  loc_00DF5931: push edx
  loc_00DF5932: push 00FFE7CEh
  loc_00DF5937: mov eax, Me
  loc_00DF593A: mov ecx, [eax]
  loc_00DF593C: mov edx, Me
  loc_00DF593F: push edx
  loc_00DF5940: call [ecx+0000090Ch]
  loc_00DF5946: mov eax, var_3C
  loc_00DF5949: push eax
  loc_00DF594A: push 00000001h
  loc_00DF594C: mov ecx, Me
  loc_00DF594F: mov edx, [ecx+000001D4h]
  loc_00DF5955: sub edx, 00000002h
  loc_00DF5958: jo 00DF5F00h
  loc_00DF595E: push edx
  loc_00DF595F: push 00000001h
  loc_00DF5961: push 00000001h
  loc_00DF5963: mov eax, Me
  loc_00DF5966: mov ecx, [eax]
  loc_00DF5968: mov edx, Me
  loc_00DF596B: push edx
  loc_00DF596C: call [ecx+000008E4h]
  loc_00DF5972: mov var_4, 0000005Fh
  loc_00DF5979: mov var_38, 00000000h
  loc_00DF5980: lea eax, var_3C
  loc_00DF5983: push eax
  loc_00DF5984: lea ecx, var_38
  loc_00DF5987: push ecx
  loc_00DF5988: push 00E6AF8Eh
  loc_00DF598D: mov edx, Me
  loc_00DF5990: mov eax, [edx]
  loc_00DF5992: mov ecx, Me
  loc_00DF5995: push ecx
  loc_00DF5996: call [eax+0000090Ch]
  loc_00DF599C: mov edx, var_3C
  loc_00DF599F: push edx
  loc_00DF59A0: mov eax, Me
  loc_00DF59A3: mov ecx, [eax+000001D0h]
  loc_00DF59A9: sub ecx, 00000002h
  loc_00DF59AC: jo 00DF5F00h
  loc_00DF59B2: push ecx
  loc_00DF59B3: push 00000001h
  loc_00DF59B5: push 00000001h
  loc_00DF59B7: push 00000001h
  loc_00DF59B9: mov edx, Me
  loc_00DF59BC: mov eax, [edx]
  loc_00DF59BE: mov ecx, Me
  loc_00DF59C1: push ecx
  loc_00DF59C2: call [eax+000008E4h]
  loc_00DF59C8: mov var_4, 00000060h
  loc_00DF59CF: mov var_38, 00000000h
  loc_00DF59D6: lea edx, var_3C
  loc_00DF59D9: push edx
  loc_00DF59DA: lea eax, var_38
  loc_00DF59DD: push eax
  loc_00DF59DE: push 00E6AF8Eh
  loc_00DF59E3: mov ecx, Me
  loc_00DF59E6: mov edx, [ecx]
  loc_00DF59E8: mov eax, Me
  loc_00DF59EB: push eax
  loc_00DF59EC: call [edx+0000090Ch]
  loc_00DF59F2: mov ecx, var_3C
  loc_00DF59F5: push ecx
  loc_00DF59F6: mov edx, Me
  loc_00DF59F9: mov eax, [edx+000001D0h]
  loc_00DF59FF: sub eax, 00000002h
  loc_00DF5A02: jo 00DF5F00h
  loc_00DF5A08: push eax
  loc_00DF5A09: mov ecx, Me
  loc_00DF5A0C: mov edx, [ecx+000001D4h]
  loc_00DF5A12: sub edx, 00000002h
  loc_00DF5A15: jo 00DF5F00h
  loc_00DF5A1B: push edx
  loc_00DF5A1C: push 00000002h
  loc_00DF5A1E: mov eax, Me
  loc_00DF5A21: mov ecx, [eax+000001D4h]
  loc_00DF5A27: sub ecx, 00000002h
  loc_00DF5A2A: jo 00DF5F00h
  loc_00DF5A30: push ecx
  loc_00DF5A31: mov edx, Me
  loc_00DF5A34: mov eax, [edx]
  loc_00DF5A36: mov ecx, Me
  loc_00DF5A39: push ecx
  loc_00DF5A3A: call [eax+000008E4h]
  loc_00DF5A40: mov var_4, 00000061h
  loc_00DF5A47: mov var_38, 00000000h
  loc_00DF5A4E: lea edx, var_3C
  loc_00DF5A51: push edx
  loc_00DF5A52: lea eax, var_38
  loc_00DF5A55: push eax
  loc_00DF5A56: push 00FFFFFFh
  loc_00DF5A5B: mov ecx, Me
  loc_00DF5A5E: mov edx, [ecx]
  loc_00DF5A60: mov eax, Me
  loc_00DF5A63: push eax
  loc_00DF5A64: call [edx+0000090Ch]
  loc_00DF5A6A: mov ecx, var_3C
  loc_00DF5A6D: push ecx
  loc_00DF5A6E: mov edx, Me
  loc_00DF5A71: mov eax, [edx+000001D0h]
  loc_00DF5A77: sub eax, 00000003h
  loc_00DF5A7A: jo 00DF5F00h
  loc_00DF5A80: push eax
  loc_00DF5A81: push 00000002h
  loc_00DF5A83: push 00000002h
  loc_00DF5A85: push 00000002h
  loc_00DF5A87: mov ecx, Me
  loc_00DF5A8A: mov edx, [ecx]
  loc_00DF5A8C: mov eax, Me
  loc_00DF5A8F: push eax
  loc_00DF5A90: call [edx+000008E4h]
  loc_00DF5A96: mov var_4, 00000062h
  loc_00DF5A9D: mov var_38, 00000000h
  loc_00DF5AA4: lea ecx, var_3C
  loc_00DF5AA7: push ecx
  loc_00DF5AA8: lea edx, var_38
  loc_00DF5AAB: push edx
  loc_00DF5AAC: push 00FFFFFFh
  loc_00DF5AB1: mov eax, Me
  loc_00DF5AB4: mov ecx, [eax]
  loc_00DF5AB6: mov edx, Me
  loc_00DF5AB9: push edx
  loc_00DF5ABA: call [ecx+0000090Ch]
  loc_00DF5AC0: mov eax, var_3C
  loc_00DF5AC3: push eax
  loc_00DF5AC4: mov ecx, Me
  loc_00DF5AC7: mov edx, [ecx+000001D0h]
  loc_00DF5ACD: sub edx, 00000003h
  loc_00DF5AD0: jo 00DF5F00h
  loc_00DF5AD6: push edx
  loc_00DF5AD7: mov eax, Me
  loc_00DF5ADA: mov ecx, [eax+000001D4h]
  loc_00DF5AE0: sub ecx, 00000003h
  loc_00DF5AE3: jo 00DF5F00h
  loc_00DF5AE9: push ecx
  loc_00DF5AEA: push 00000003h
  loc_00DF5AEC: mov edx, Me
  loc_00DF5AEF: mov eax, [edx+000001D4h]
  loc_00DF5AF5: sub eax, 00000003h
  loc_00DF5AF8: jo 00DF5F00h
  loc_00DF5AFE: push eax
  loc_00DF5AFF: mov ecx, Me
  loc_00DF5B02: mov edx, [ecx]
  loc_00DF5B04: mov eax, Me
  loc_00DF5B07: push eax
  loc_00DF5B08: call [edx+000008E4h]
  loc_00DF5B0E: mov var_4, 00000063h
  loc_00DF5B15: mov var_38, 00000000h
  loc_00DF5B1C: lea ecx, var_3C
  loc_00DF5B1F: push ecx
  loc_00DF5B20: lea edx, var_38
  loc_00DF5B23: push edx
  loc_00DF5B24: push 00E4AD89h
  loc_00DF5B29: mov eax, Me
  loc_00DF5B2C: mov ecx, [eax]
  loc_00DF5B2E: mov edx, Me
  loc_00DF5B31: push edx
  loc_00DF5B32: call [ecx+0000090Ch]
  loc_00DF5B38: mov eax, var_3C
  loc_00DF5B3B: push eax
  loc_00DF5B3C: mov ecx, Me
  loc_00DF5B3F: mov edx, [ecx+000001D0h]
  loc_00DF5B45: sub edx, 00000003h
  loc_00DF5B48: jo 00DF5F00h
  loc_00DF5B4E: push edx
  loc_00DF5B4F: mov eax, Me
  loc_00DF5B52: mov ecx, [eax+000001D4h]
  loc_00DF5B58: sub ecx, 00000002h
  loc_00DF5B5B: jo 00DF5F00h
  loc_00DF5B61: push ecx
  loc_00DF5B62: mov edx, Me
  loc_00DF5B65: mov eax, [edx+000001D0h]
  loc_00DF5B6B: sub eax, 00000003h
  loc_00DF5B6E: jo 00DF5F00h
  loc_00DF5B74: push eax
  loc_00DF5B75: push 00000002h
  loc_00DF5B77: mov ecx, Me
  loc_00DF5B7A: mov edx, [ecx]
  loc_00DF5B7C: mov eax, Me
  loc_00DF5B7F: push eax
  loc_00DF5B80: call [edx+000008E4h]
  loc_00DF5B86: mov var_4, 00000064h
  loc_00DF5B8D: mov var_38, 00000000h
  loc_00DF5B94: lea ecx, var_3C
  loc_00DF5B97: push ecx
  loc_00DF5B98: lea edx, var_38
  loc_00DF5B9B: push edx
  loc_00DF5B9C: push 00EE8269h
  loc_00DF5BA1: mov eax, Me
  loc_00DF5BA4: mov ecx, [eax]
  loc_00DF5BA6: mov edx, Me
  loc_00DF5BA9: push edx
  loc_00DF5BAA: call [ecx+0000090Ch]
  loc_00DF5BB0: mov eax, var_3C
  loc_00DF5BB3: push eax
  loc_00DF5BB4: mov ecx, Me
  loc_00DF5BB7: mov edx, [ecx+000001D0h]
  loc_00DF5BBD: sub edx, 00000002h
  loc_00DF5BC0: jo 00DF5F00h
  loc_00DF5BC6: push edx
  loc_00DF5BC7: mov eax, Me
  loc_00DF5BCA: mov ecx, [eax+000001D4h]
  loc_00DF5BD0: sub ecx, 00000002h
  loc_00DF5BD3: jo 00DF5F00h
  loc_00DF5BD9: push ecx
  loc_00DF5BDA: mov edx, Me
  loc_00DF5BDD: mov eax, [edx+000001D0h]
  loc_00DF5BE3: sub eax, 00000002h
  loc_00DF5BE6: jo 00DF5F00h
  loc_00DF5BEC: push eax
  loc_00DF5BED: push 00000002h
  loc_00DF5BEF: mov ecx, Me
  loc_00DF5BF2: mov edx, [ecx]
  loc_00DF5BF4: mov eax, Me
  loc_00DF5BF7: push eax
  loc_00DF5BF8: call [edx+000008E4h]
  loc_00DF5BFE: mov var_4, 00000068h
  loc_00DF5C05: push FFFFFFFFh
  loc_00DF5C07: call [004010A4h] ; __vbaOnError
  loc_00DF5C0D: mov var_4, 00000069h
  loc_00DF5C14: mov ecx, Me
  loc_00DF5C17: movsx edx, [ecx+00000070h]
  loc_00DF5C1B: test edx, edx
  loc_00DF5C1D: jz 00DF5D04h
  loc_00DF5C23: mov var_4, 0000006Ah
  loc_00DF5C2A: mov eax, Me
  loc_00DF5C2D: movsx ecx, [eax+0000006Eh]
  loc_00DF5C31: test ecx, ecx
  loc_00DF5C33: jz 00DF5D04h
  loc_00DF5C39: mov edx, Me
  loc_00DF5C3C: movsx eax, [edx+00000070h]
  loc_00DF5C40: test eax, eax
  loc_00DF5C42: jz 00DF5D04h
  loc_00DF5C48: mov ecx, Me
  loc_00DF5C4B: movsx edx, [ecx+00000050h]
  loc_00DF5C4F: test edx, edx
  loc_00DF5C51: jnz 00DF5C62h
  loc_00DF5C53: mov eax, Me
  loc_00DF5C56: movsx ecx, [eax+00000058h]
  loc_00DF5C5A: test ecx, ecx
  loc_00DF5C5C: jz 00DF5D04h
  loc_00DF5C62: mov var_4, 0000006Bh
  loc_00DF5C69: mov edx, Me
  loc_00DF5C6C: mov eax, [edx+000001D0h]
  loc_00DF5C72: sub eax, 00000002h
  loc_00DF5C75: jo 00DF5F00h
  loc_00DF5C7B: push eax
  loc_00DF5C7C: mov ecx, Me
  loc_00DF5C7F: mov edx, [ecx+000001D4h]
  loc_00DF5C85: sub edx, 00000002h
  loc_00DF5C88: jo 00DF5F00h
  loc_00DF5C8E: push edx
  loc_00DF5C8F: push 00000002h
  loc_00DF5C91: push 00000002h
  loc_00DF5C93: lea eax, var_30
  loc_00DF5C96: push eax
  loc_00DF5C97: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF5C9C: call [00401074h] ; __vbaSetSystemError
  loc_00DF5CA2: mov var_4, 0000006Ch
  loc_00DF5CA9: lea ecx, var_38
  loc_00DF5CAC: push ecx
  loc_00DF5CAD: mov edx, Me
  loc_00DF5CB0: mov eax, [edx]
  loc_00DF5CB2: mov ecx, Me
  loc_00DF5CB5: push ecx
  loc_00DF5CB6: call [eax+000000D8h]
  loc_00DF5CBC: fnclex
  loc_00DF5CBE: mov var_50, eax
  loc_00DF5CC1: cmp var_50, 00000000h
  loc_00DF5CC5: jge 00DF5CE7h
  loc_00DF5CC7: push 000000D8h
  loc_00DF5CCC: push 006DCB00h
  loc_00DF5CD1: mov edx, Me
  loc_00DF5CD4: push edx
  loc_00DF5CD5: mov eax, var_50
  loc_00DF5CD8: push eax
  loc_00DF5CD9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF5CDF: mov var_CC, eax
  loc_00DF5CE5: jmp 00DF5CF1h
  loc_00DF5CE7: mov var_CC, 00000000h
  loc_00DF5CF1: lea ecx, var_30
  loc_00DF5CF4: push ecx
  loc_00DF5CF5: mov edx, var_38
  loc_00DF5CF8: push edx
  loc_00DF5CF9: call 006DDA18h ; DrawFocusRect(%x1v, %x2v)
  loc_00DF5CFE: call [00401074h] ; __vbaSetSystemError
  loc_00DF5D04: mov var_4, 0000006Fh
  loc_00DF5D0B: mov eax, Me
  loc_00DF5D0E: mov ecx, [eax+000000CCh]
  loc_00DF5D14: mov var_68, ecx
  loc_00DF5D17: mov edx, var_68
  loc_00DF5D1A: mov var_D0, edx
  loc_00DF5D20: cmp var_D0, 00000003h
  loc_00DF5D27: ja 00DF5E96h
  loc_00DF5D2D: mov eax, var_D0
  loc_00DF5D33: jmp [eax*4+00DF5EEBh]
  loc_00DF5D3A: jmp 00DF5E96h
  loc_00DF5D3F: mov var_4, 00000071h
  loc_00DF5D46: mov var_38, 00000000h
  loc_00DF5D4D: lea ecx, var_3C
  loc_00DF5D50: push ecx
  loc_00DF5D51: lea edx, var_38
  loc_00DF5D54: push edx
  loc_00DF5D55: push 00743C00h
  loc_00DF5D5A: mov eax, Me
  loc_00DF5D5D: mov ecx, [eax]
  loc_00DF5D5F: mov edx, Me
  loc_00DF5D62: push edx
  loc_00DF5D63: call [ecx+0000090Ch]
  loc_00DF5D69: mov eax, var_3C
  loc_00DF5D6C: push eax
  loc_00DF5D6D: mov ecx, Me
  loc_00DF5D70: mov edx, [ecx+000001D0h]
  loc_00DF5D76: push edx
  loc_00DF5D77: mov eax, Me
  loc_00DF5D7A: mov ecx, [eax+000001D4h]
  loc_00DF5D80: push ecx
  loc_00DF5D81: push 00000000h
  loc_00DF5D83: push 00000000h
  loc_00DF5D85: mov edx, Me
  loc_00DF5D88: mov eax, [edx]
  loc_00DF5D8A: mov ecx, Me
  loc_00DF5D8D: push ecx
  loc_00DF5D8E: call [eax+000008F0h]
  loc_00DF5D94: mov var_4, 00000072h
  loc_00DF5D9B: mov var_38, 00000000h
  loc_00DF5DA2: lea edx, var_3C
  loc_00DF5DA5: push edx
  loc_00DF5DA6: lea eax, var_38
  loc_00DF5DA9: push eax
  loc_00DF5DAA: push 00743C00h
  loc_00DF5DAF: mov ecx, Me
  loc_00DF5DB2: mov edx, [ecx]
  loc_00DF5DB4: mov eax, Me
  loc_00DF5DB7: push eax
  loc_00DF5DB8: call [edx+0000090Ch]
  loc_00DF5DBE: mov var_44, 3E99999Ah
  loc_00DF5DC5: mov ecx, var_3C
  loc_00DF5DC8: mov var_40, ecx
  loc_00DF5DCB: lea edx, var_48
  loc_00DF5DCE: push edx
  loc_00DF5DCF: lea eax, var_44
  loc_00DF5DD2: push eax
  loc_00DF5DD3: lea ecx, var_40
  loc_00DF5DD6: push ecx
  loc_00DF5DD7: mov edx, Me
  loc_00DF5DDA: mov eax, [edx]
  loc_00DF5DDC: mov ecx, Me
  loc_00DF5DDF: push ecx
  loc_00DF5DE0: call [eax+0000097Ch]
  loc_00DF5DE6: mov edx, var_48
  loc_00DF5DE9: mov var_4C, edx
  loc_00DF5DEC: lea eax, var_4C
  loc_00DF5DEF: push eax
  loc_00DF5DF0: mov ecx, Me
  loc_00DF5DF3: mov edx, [ecx]
  loc_00DF5DF5: mov eax, Me
  loc_00DF5DF8: push eax
  loc_00DF5DF9: call [edx+00000940h]
  loc_00DF5DFF: jmp 00DF5E96h
  loc_00DF5E04: mov var_4, 00000074h
  loc_00DF5E0B: push 00000006h
  loc_00DF5E0D: push 00000062h
  loc_00DF5E0F: push 00000037h
  loc_00DF5E11: call [00401028h] ; rtcRgb
  loc_00DF5E17: push eax
  loc_00DF5E18: mov ecx, Me
  loc_00DF5E1B: mov edx, [ecx+000001D0h]
  loc_00DF5E21: push edx
  loc_00DF5E22: mov eax, Me
  loc_00DF5E25: mov ecx, [eax+000001D4h]
  loc_00DF5E2B: push ecx
  loc_00DF5E2C: push 00000000h
  loc_00DF5E2E: push 00000000h
  loc_00DF5E30: mov edx, Me
  loc_00DF5E33: mov eax, [edx]
  loc_00DF5E35: mov ecx, Me
  loc_00DF5E38: push ecx
  loc_00DF5E39: call [eax+000008F0h]
  loc_00DF5E3F: mov var_4, 00000075h
  loc_00DF5E46: push 00000006h
  loc_00DF5E48: push 00000062h
  loc_00DF5E4A: push 00000037h
  loc_00DF5E4C: call [00401028h] ; rtcRgb
  loc_00DF5E52: mov var_48, eax
  loc_00DF5E55: mov var_3C, 3E99999Ah
  loc_00DF5E5C: mov edx, var_48
  loc_00DF5E5F: mov var_38, edx
  loc_00DF5E62: lea eax, var_40
  loc_00DF5E65: push eax
  loc_00DF5E66: lea ecx, var_3C
  loc_00DF5E69: push ecx
  loc_00DF5E6A: lea edx, var_38
  loc_00DF5E6D: push edx
  loc_00DF5E6E: mov eax, Me
  loc_00DF5E71: mov ecx, [eax]
  loc_00DF5E73: mov edx, Me
  loc_00DF5E76: push edx
  loc_00DF5E77: call [ecx+0000097Ch]
  loc_00DF5E7D: mov eax, var_40
  loc_00DF5E80: mov var_44, eax
  loc_00DF5E83: lea ecx, var_44
  loc_00DF5E86: push ecx
  loc_00DF5E87: mov edx, Me
  loc_00DF5E8A: mov eax, [edx]
  loc_00DF5E8C: mov ecx, Me
  loc_00DF5E8F: push ecx
  loc_00DF5E90: call [eax+00000940h]
  loc_00DF5E96: xor eax, eax
  loc_00DF5E98: mov ecx, var_20
  loc_00DF5E9B: mov fs:[00000000h], ecx
  loc_00DF5EA2: pop edi
  loc_00DF5EA3: pop esi
  loc_00DF5EA4: pop ebx
  loc_00DF5EA5: mov esp, ebp
  loc_00DF5EA7: pop ebp
  loc_00DF5EA8: retn 0008h
End Sub

Private Sub Proc_2_127_DF5F10(arg_C, arg_10, arg_14, arg_18, arg_1C) 'DF5F10
  loc_00DF5F10: sub esp, 00000020h
  loc_00DF5F13: xor eax, eax
  loc_00DF5F15: push ebx
  loc_00DF5F16: mov arg_C, eax
  loc_00DF5F1A: push ebp
  loc_00DF5F1B: mov arg_14, eax
  loc_00DF5F1F: push esi
  loc_00DF5F20: mov esi, arg_28
  loc_00DF5F24: mov arg_1C, eax
  loc_00DF5F28: push edi
  loc_00DF5F29: mov arg_24, eax
  loc_00DF5F2D: mov eax, [esi+00000010h]
  loc_00DF5F30: xor edi, edi
  loc_00DF5F32: lea edx, Me
  loc_00DF5F36: mov arg_10, edi
  loc_00DF5F3A: mov arg_14, edi
  loc_00DF5F3E: mov Me, edi
  loc_00DF5F42: mov arg_C, edi
  loc_00DF5F46: mov ecx, [eax]
  loc_00DF5F48: push edx
  loc_00DF5F49: push eax
  loc_00DF5F4A: xor ebx, ebx
  loc_00DF5F4C: call [ecx+00000108h]
  loc_00DF5F52: cmp eax, edi
  loc_00DF5F54: fnclex
  loc_00DF5F56: jge 00DF5F6Dh
  loc_00DF5F58: mov ecx, [esi+00000010h]
  loc_00DF5F5B: push 00000108h
  loc_00DF5F60: push 006DCB00h
  loc_00DF5F65: push ecx
  loc_00DF5F66: push eax
  loc_00DF5F67: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF5F6D: fld real4 ptr Me
  loc_00DF5F71: mov ebp, [00401208h] ; __vbaFpI4
  loc_00DF5F77: call ebp
  loc_00DF5F79: mov [esi+000001D0h], eax
  loc_00DF5F7F: mov eax, [esi+00000010h]
  loc_00DF5F82: lea ecx, Me
  loc_00DF5F86: mov edx, [eax]
  loc_00DF5F88: push ecx
  loc_00DF5F89: push eax
  loc_00DF5F8A: call [edx+00000100h]
  loc_00DF5F90: cmp eax, edi
  loc_00DF5F92: fnclex
  loc_00DF5F94: jge 00DF5FABh
  loc_00DF5F96: mov edx, [esi+00000010h]
  loc_00DF5F99: push 00000100h
  loc_00DF5F9E: push 006DCB00h
  loc_00DF5FA3: push edx
  loc_00DF5FA4: push eax
  loc_00DF5FA5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF5FAB: fld real4 ptr Me
  loc_00DF5FAF: call ebp
  loc_00DF5FB1: lea ecx, arg_C
  loc_00DF5FB5: lea edx, Me
  loc_00DF5FB9: push ecx
  loc_00DF5FBA: mov ecx, [esi+00000088h]
  loc_00DF5FC0: mov [esi+000001D4h], eax
  loc_00DF5FC6: mov eax, [esi]
  loc_00DF5FC8: push edx
  loc_00DF5FC9: push ecx
  loc_00DF5FCA: push esi
  loc_00DF5FCB: mov arg_18, edi
  loc_00DF5FCF: call [eax+0000090Ch]
  loc_00DF5FD5: mov eax, [esi+000001D0h]
  loc_00DF5FDB: mov ecx, [esi+000001D4h]
  loc_00DF5FE1: mov edx, arg_C
  loc_00DF5FE5: push eax
  loc_00DF5FE6: mov arg_18, edx
  loc_00DF5FEA: push ecx
  loc_00DF5FEB: push edi
  loc_00DF5FEC: lea edx, arg_24
  loc_00DF5FF0: push edi
  loc_00DF5FF1: push edx
  loc_00DF5FF2: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF5FF7: call [00401074h] ; __vbaSetSystemError
  loc_00DF5FFD: mov eax, [esi+000000CCh]
  loc_00DF6003: cmp eax, 00000003h
  loc_00DF6006: ja 00DF6114h
  loc_00DF600C: jmp [eax*4+00DF62BCh]
  loc_00DF6013: mov eax, [esi]
  loc_00DF6015: lea ecx, arg_C
  loc_00DF6019: lea edx, Me
  loc_00DF601D: push ecx
  loc_00DF601E: push edx
  loc_00DF601F: push 00EED2C1h
  loc_00DF6024: push esi
  loc_00DF6025: mov arg_18, edi
  loc_00DF6029: call [eax+0000090Ch]
  loc_00DF602F: mov ecx, [esi]
  loc_00DF6031: lea edx, arg_C
  loc_00DF6035: mov eax, arg_C
  loc_00DF6039: push edx
  loc_00DF603A: mov arg_14, eax
  loc_00DF603E: lea eax, arg_C
  loc_00DF6042: push eax
  loc_00DF6043: push 00C56A31h
  loc_00DF6048: push esi
  loc_00DF6049: mov arg_18, edi
  loc_00DF604D: call [ecx+0000090Ch]
  loc_00DF6053: jmp 00DF6110h
  loc_00DF6058: mov ecx, [esi]
  loc_00DF605A: lea edx, arg_C
  loc_00DF605E: lea eax, Me
  loc_00DF6062: push edx
  loc_00DF6063: push eax
  loc_00DF6064: push 00E3DFE0h
  loc_00DF6069: push esi
  loc_00DF606A: mov arg_18, edi
  loc_00DF606E: call [ecx+0000090Ch]
  loc_00DF6074: mov edx, [esi]
  loc_00DF6076: lea eax, arg_C
  loc_00DF607A: mov ecx, arg_C
  loc_00DF607E: push eax
  loc_00DF607F: mov arg_14, ecx
  loc_00DF6083: lea ecx, arg_C
  loc_00DF6087: push ecx
  loc_00DF6088: push 00BFB4B2h
  loc_00DF608D: push esi
  loc_00DF608E: mov arg_18, edi
  loc_00DF6092: call [edx+0000090Ch]
  loc_00DF6098: jmp 00DF6110h
  loc_00DF609A: mov edx, [esi]
  loc_00DF609C: lea eax, arg_C
  loc_00DF60A0: lea ecx, Me
  loc_00DF60A4: push eax
  loc_00DF60A5: push ecx
  loc_00DF60A6: push 00BAD6D4h
  loc_00DF60AB: push esi
  loc_00DF60AC: mov arg_18, edi
  loc_00DF60B0: call [edx+0000090Ch]
  loc_00DF60B6: mov eax, [esi]
  loc_00DF60B8: lea ecx, arg_C
  loc_00DF60BC: mov edx, arg_C
  loc_00DF60C0: push ecx
  loc_00DF60C1: mov arg_14, edx
  loc_00DF60C5: lea edx, arg_C
  loc_00DF60C9: push edx
  loc_00DF60CA: push 0070A093h
  loc_00DF60CF: push esi
  loc_00DF60D0: mov arg_18, edi
  loc_00DF60D4: call [eax+0000090Ch]
  loc_00DF60DA: jmp 00DF6110h
  loc_00DF60DC: mov eax, arg_14
  loc_00DF60E0: mov ecx, [esi]
  loc_00DF60E2: fld real8 ptr [004017F8h]
  loc_00DF60E8: mov arg_10, eax
  loc_00DF60EC: lea edx, arg_C
  loc_00DF60F0: fstp real4 ptr Me
  loc_00DF60F4: fnstsw ax
  loc_00DF60F6: test al, 0Dh
  loc_00DF60F8: jnz 00DF62CCh
  loc_00DF60FE: push edx
  loc_00DF60FF: lea eax, arg_C
  loc_00DF6103: lea edx, arg_18
  loc_00DF6107: push eax
  loc_00DF6108: push edx
  loc_00DF6109: push esi
  loc_00DF610A: call [ecx+0000097Ch]
  loc_00DF6110: mov ebx, arg_C
  loc_00DF6114: cmp [esi+00000064h], edi
  loc_00DF6117: jz 00DF61F2h
  loc_00DF611D: cmp [esi+0000006Ch], di
  loc_00DF6121: jz 00DF61F2h
  loc_00DF6127: fld real8 ptr [00401590h]
  loc_00DF612D: fstp real4 ptr Me
  loc_00DF6131: fnstsw ax
  loc_00DF6133: test al, 0Dh
  loc_00DF6135: jnz 00DF62CCh
  loc_00DF613B: lea ecx, arg_C
  loc_00DF613F: lea edx, Me
  loc_00DF6143: push ecx
  loc_00DF6144: lea ecx, arg_14
  loc_00DF6148: mov eax, [esi]
  loc_00DF614A: push edx
  loc_00DF614B: push ecx
  loc_00DF614C: push esi
  loc_00DF614D: call [eax+0000097Ch]
  loc_00DF6153: mov edx, [esi]
  loc_00DF6155: lea ebp, [esi+000000ACh]
  loc_00DF615B: mov eax, arg_C
  loc_00DF615F: push ebp
  loc_00DF6160: push eax
  loc_00DF6161: push esi
  loc_00DF6162: call [edx+00000974h]
  loc_00DF6168: mov edx, [esi+000001D0h]
  loc_00DF616E: mov eax, [esi+000001D4h]
  loc_00DF6174: mov ecx, [esi]
  loc_00DF6176: push ebx
  loc_00DF6177: push edx
  loc_00DF6178: push eax
  loc_00DF6179: push edi
  loc_00DF617A: push edi
  loc_00DF617B: push esi
  loc_00DF617C: call [ecx+000008F0h]
  loc_00DF6182: cmp [esi+0000004Eh], di
  loc_00DF6186: jz 00DF61DDh
  loc_00DF6188: fld real8 ptr [004017F0h]
  loc_00DF618E: fstp real4 ptr Me
  loc_00DF6192: fnstsw ax
  loc_00DF6194: test al, 0Dh
  loc_00DF6196: jnz 00DF62CCh
  loc_00DF619C: mov ecx, [esi]
  loc_00DF619E: lea edx, arg_C
  loc_00DF61A2: push edx
  loc_00DF61A3: lea eax, arg_C
  loc_00DF61A7: lea edx, arg_14
  loc_00DF61AB: push eax
  loc_00DF61AC: push edx
  loc_00DF61AD: push esi
  loc_00DF61AE: call [ecx+0000097Ch]
  loc_00DF61B4: mov eax, [esi]
  loc_00DF61B6: push ebp
  loc_00DF61B7: mov ecx, arg_10
  loc_00DF61BB: push ecx
  loc_00DF61BC: push esi
  loc_00DF61BD: call [eax+00000974h]
  loc_00DF61C3: mov eax, [esi+000001D0h]
  loc_00DF61C9: mov ecx, [esi+000001D4h]
  loc_00DF61CF: mov edx, [esi]
  loc_00DF61D1: push ebx
  loc_00DF61D2: push eax
  loc_00DF61D3: push ecx
  loc_00DF61D4: push edi
  loc_00DF61D5: push edi
  loc_00DF61D6: push esi
  loc_00DF61D7: call [edx+000008F0h]
  loc_00DF61DD: mov edx, [esi]
  loc_00DF61DF: push esi
  loc_00DF61E0: call [edx+00000920h]
  loc_00DF61E6: pop edi
  loc_00DF61E7: pop esi
  loc_00DF61E8: pop ebp
  loc_00DF61E9: xor eax, eax
  loc_00DF61EB: pop ebx
  loc_00DF61EC: add esp, 00000020h
  loc_00DF61EF: retn 0008h
End Sub

Private Sub Proc_2_128_DF62E0(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28, arg_2C, arg_30, arg_34) 'DF62E0
  loc_00DF62E0: sub esp, 00000040h
  loc_00DF62E3: push ebx
  loc_00DF62E4: xor eax, eax
  loc_00DF62E6: push ebp
  loc_00DF62E7: push esi
  loc_00DF62E8: mov esi, Proc_2_128_DF62E08
  loc_00DF62EC: mov arg_34, eax
  loc_00DF62F0: mov arg_38, eax
  loc_00DF62F4: lea edx, Proc_2_128_DF62E0
  loc_00DF62F8: mov ecx, [esi]
  loc_00DF62FA: push edi
  loc_00DF62FB: mov Proc_2_128_DF62E00, eax
  loc_00DF62FF: xor edi, edi
  loc_00DF6301: push edx
  loc_00DF6302: push esi
  loc_00DF6303: mov Proc_2_128_DF62E0C, eax
  loc_00DF6307: mov arg_10, edi
  loc_00DF630B: mov arg_14, edi
  loc_00DF630F: mov arg_1C, edi
  loc_00DF6313: mov arg_18, edi
  loc_00DF6317: mov arg_20, edi
  loc_00DF631B: mov arg_24, edi
  loc_00DF631F: mov arg_28, edi
  loc_00DF6323: mov arg_30, edi
  loc_00DF6327: mov arg_2C, edi
  loc_00DF632B: mov arg_34, edi
  loc_00DF632F: call [ecx+00000108h]
  loc_00DF6335: mov ebp, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF633B: cmp eax, edi
  loc_00DF633D: fnclex
  loc_00DF633F: jge 00DF634Fh
  loc_00DF6341: push 00000108h
  loc_00DF6346: push 006DCB00h
  loc_00DF634B: push esi
  loc_00DF634C: push eax
  loc_00DF634D: call ebp
  loc_00DF634F: fld real4 ptr Me
  loc_00DF6353: mov ebx, [00401208h] ; __vbaFpI4
  loc_00DF6359: call ebx
  loc_00DF635B: lea ecx, Me
  loc_00DF635F: mov [esi+000001D0h], eax
  loc_00DF6365: mov eax, [esi]
  loc_00DF6367: push ecx
  loc_00DF6368: push esi
  loc_00DF6369: call [eax+00000100h]
  loc_00DF636F: cmp eax, edi
  loc_00DF6371: fnclex
  loc_00DF6373: jge 00DF6383h
  loc_00DF6375: push 00000100h
  loc_00DF637A: push 006DCB00h
  loc_00DF637F: push esi
  loc_00DF6380: push eax
  loc_00DF6381: call ebp
  loc_00DF6383: fld real4 ptr Me
  loc_00DF6387: call ebx
  loc_00DF6389: cmp [esi+0000007Ch], di
  loc_00DF638D: mov [esi+000001D4h], eax
  loc_00DF6393: jnz 00DF6399h
  loc_00DF6395: xor eax, eax
  loc_00DF6397: jmp 00DF639Dh
  loc_00DF6399: mov eax, arg_50
  loc_00DF639D: sub eax, edi
  loc_00DF639F: jz 00DF6985h
  loc_00DF63A5: dec eax
  loc_00DF63A6: jz 00DF673Ch
  loc_00DF63AC: dec eax
  loc_00DF63AD: jnz 00DF6C30h
  loc_00DF63B3: mov edx, [esi]
  loc_00DF63B5: lea eax, arg_C
  loc_00DF63B9: lea ecx, Me
  loc_00DF63BD: push eax
  loc_00DF63BE: push ecx
  loc_00DF63BF: push 00FFFFFFh
  loc_00DF63C4: push esi
  loc_00DF63C5: mov arg_18, edi
  loc_00DF63C9: call [edx+0000090Ch]
  loc_00DF63CF: mov edx, [esi]
  loc_00DF63D1: lea eax, arg_10
  loc_00DF63D5: push eax
  loc_00DF63D6: mov eax, [esi+00000088h]
  loc_00DF63DC: lea ecx, arg_18
  loc_00DF63E0: mov arg_18, edi
  loc_00DF63E4: push ecx
  loc_00DF63E5: push eax
  loc_00DF63E6: push esi
  loc_00DF63E7: call [edx+0000090Ch]
  loc_00DF63ED: fld real8 ptr [004013B0h]
  loc_00DF63F3: mov ecx, arg_10
  loc_00DF63F7: mov edx, [esi]
  loc_00DF63F9: fstp real4 ptr arg_1C
  loc_00DF63FD: fnstsw ax
  loc_00DF63FF: test al, 0Dh
  loc_00DF6401: jnz 00DF6D15h
  loc_00DF6407: lea eax, arg_20
  loc_00DF640B: mov arg_18, ecx
  loc_00DF640F: push eax
  loc_00DF6410: lea ecx, arg_20
  loc_00DF6414: lea eax, arg_1C
  loc_00DF6418: push ecx
  loc_00DF6419: push eax
  loc_00DF641A: push esi
  loc_00DF641B: call [edx+0000097Ch]
  loc_00DF6421: fild real4 ptr [esi+000001D0h]
  loc_00DF6427: mov ecx, arg_20
  loc_00DF642B: mov edx, arg_C
  loc_00DF642F: mov ebp, [esi]
  loc_00DF6431: push 00000001h
  loc_00DF6433: fstp real8 ptr arg_34
  loc_00DF6437: fld real8 ptr arg_34
  loc_00DF643B: cmp [00E3D000h], 00000000h
  loc_00DF6442: jnz 00DF644Ch
  loc_00DF6444: fdiv st0, real8 ptr [00401338h]
  loc_00DF644A: jmp 00DF645Dh
  loc_00DF644C: push [0040133Ch]
  loc_00DF6452: push [00401338h]
  loc_00DF6458: call 00402854h ; _adj_fdiv_m64
  loc_00DF645D: push ecx
  loc_00DF645E: push edx
  loc_00DF645F: fnstsw ax
  loc_00DF6461: test al, 0Dh
  loc_00DF6463: jnz 00DF6D15h
  loc_00DF6469: call ebx
  loc_00DF646B: push eax
  loc_00DF646C: mov eax, [esi+000001D4h]
  loc_00DF6472: push eax
  loc_00DF6473: push edi
  loc_00DF6474: push edi
  loc_00DF6475: push esi
  loc_00DF6476: call arg_908
  loc_00DF647C: mov ecx, [esi]
  loc_00DF647E: lea edx, arg_C
  loc_00DF6482: push edx
  loc_00DF6483: mov edx, [esi+00000088h]
  loc_00DF6489: lea eax, arg_C
  loc_00DF648D: mov arg_C, edi
  loc_00DF6491: push eax
  loc_00DF6492: push edx
  loc_00DF6493: push esi
  loc_00DF6494: call [ecx+0000090Ch]
  loc_00DF649A: fld real8 ptr [004013B0h]
  loc_00DF64A0: fstp real4 ptr arg_10
  loc_00DF64A4: fnstsw ax
  loc_00DF64A6: test al, 0Dh
  loc_00DF64A8: jnz 00DF6D15h
  loc_00DF64AE: mov ecx, [esi]
  loc_00DF64B0: lea edx, arg_18
  loc_00DF64B4: mov eax, arg_C
  loc_00DF64B8: push edx
  loc_00DF64B9: mov arg_18, eax
  loc_00DF64BD: lea eax, arg_14
  loc_00DF64C1: lea edx, arg_18
  loc_00DF64C5: push eax
  loc_00DF64C6: push edx
  loc_00DF64C7: push esi
  loc_00DF64C8: call [ecx+0000097Ch]
  loc_00DF64CE: mov eax, [esi]
  loc_00DF64D0: mov arg_1C, edi
  loc_00DF64D4: lea ecx, arg_20
  loc_00DF64D8: lea edx, arg_1C
  loc_00DF64DC: push ecx
  loc_00DF64DD: mov ecx, [esi+00000088h]
  loc_00DF64E3: push edx
  loc_00DF64E4: push ecx
  loc_00DF64E5: push esi
  loc_00DF64E6: call [eax+0000090Ch]
  loc_00DF64EC: fld real8 ptr [00401590h]
  loc_00DF64F2: mov edx, arg_20
  loc_00DF64F6: lea ecx, arg_2C
  loc_00DF64FA: fstp real4 ptr arg_24
  loc_00DF64FE: fnstsw ax
  loc_00DF6500: test al, 0Dh
  loc_00DF6502: jnz 00DF6D15h
  loc_00DF6508: mov arg_28, edx
  loc_00DF650C: push ecx
  loc_00DF650D: lea edx, arg_28
  loc_00DF6511: lea ecx, arg_2C
  loc_00DF6515: mov eax, [esi]
  loc_00DF6517: push edx
  loc_00DF6518: push ecx
  loc_00DF6519: push esi
  loc_00DF651A: call [eax+0000097Ch]
  loc_00DF6520: mov eax, [esi+000001D0h]
  loc_00DF6526: mov ebp, [esi]
  loc_00DF6528: mov Proc_2_128_DF62E0C, eax
  loc_00DF652C: mov edx, arg_2C
  loc_00DF6530: fild real4 ptr Proc_2_128_DF62E0C
  loc_00DF6534: mov ecx, arg_18
  loc_00DF6538: push 00000001h
  loc_00DF653A: push edx
  loc_00DF653B: mov edx, [esi+000001D4h]
  loc_00DF6541: fstp real8 ptr arg_38
  loc_00DF6545: fld real8 ptr arg_38
  loc_00DF6549: cmp [00E3D000h], 00000000h
  loc_00DF6550: jnz 00DF655Ah
  loc_00DF6552: fdiv st0, real8 ptr [00401338h]
  loc_00DF6558: jmp 00DF656Bh
  loc_00DF655A: push [0040133Ch]
  loc_00DF6560: push [00401338h]
  loc_00DF6566: call 00402854h ; _adj_fdiv_m64
  loc_00DF656B: push ecx
  loc_00DF656C: push eax
  loc_00DF656D: push edx
  loc_00DF656E: fnstsw ax
  loc_00DF6570: test al, 0Dh
  loc_00DF6572: jnz 00DF6D15h
  loc_00DF6578: call ebx
  loc_00DF657A: push eax
  loc_00DF657B: push edi
  loc_00DF657C: push esi
  loc_00DF657D: call arg_908
  loc_00DF6583: mov eax, [esi]
  loc_00DF6585: push esi
  loc_00DF6586: call [eax+00000920h]
  loc_00DF658C: mov ecx, [esi]
  loc_00DF658E: lea edx, arg_C
  loc_00DF6592: lea eax, Me
  loc_00DF6596: push edx
  loc_00DF6597: push eax
  loc_00DF6598: push 00FFFFFFh
  loc_00DF659D: push esi
  loc_00DF659E: mov arg_18, edi
  loc_00DF65A2: call [ecx+0000090Ch]
  loc_00DF65A8: mov eax, [esi+000001D0h]
  loc_00DF65AE: mov ecx, [esi]
  loc_00DF65B0: mov edx, arg_C
  loc_00DF65B4: push edx
  loc_00DF65B5: mov edx, [esi+000001D4h]
  loc_00DF65BB: sub edx, 00000002h
  loc_00DF65BE: push eax
  loc_00DF65BF: jo 00DF6D1Ah
  loc_00DF65C5: push edx
  loc_00DF65C6: push 00000001h
  loc_00DF65C8: push 00000001h
  loc_00DF65CA: push esi
  loc_00DF65CB: call [ecx+000008F0h]
  loc_00DF65D1: mov eax, [esi]
  loc_00DF65D3: lea ecx, arg_C
  loc_00DF65D7: push ecx
  loc_00DF65D8: mov ecx, [esi+00000088h]
  loc_00DF65DE: lea edx, arg_C
  loc_00DF65E2: mov arg_C, edi
  loc_00DF65E6: push edx
  loc_00DF65E7: push ecx
  loc_00DF65E8: push esi
  loc_00DF65E9: call [eax+0000090Ch]
  loc_00DF65EF: fld real8 ptr [00401810h]
  loc_00DF65F5: mov edx, arg_C
  loc_00DF65F9: lea ecx, arg_18
  loc_00DF65FD: fstp real4 ptr arg_10
  loc_00DF6601: fnstsw ax
  loc_00DF6603: test al, 0Dh
  loc_00DF6605: jnz 00DF6D15h
  loc_00DF660B: mov arg_14, edx
  loc_00DF660F: push ecx
  loc_00DF6610: lea edx, arg_14
  loc_00DF6614: lea ecx, arg_18
  loc_00DF6618: mov eax, [esi]
  loc_00DF661A: push edx
  loc_00DF661B: push ecx
  loc_00DF661C: push esi
  loc_00DF661D: call [eax+0000097Ch]
  loc_00DF6623: mov ecx, [esi+000001D0h]
  loc_00DF6629: mov edx, [esi]
  loc_00DF662B: mov eax, arg_18
  loc_00DF662F: push eax
  loc_00DF6630: mov eax, [esi+000001D4h]
  loc_00DF6636: push ecx
  loc_00DF6637: push eax
  loc_00DF6638: push edi
  loc_00DF6639: push edi
  loc_00DF663A: push esi
  loc_00DF663B: call [edx+000008F0h]
  loc_00DF6641: mov ecx, [esi]
  loc_00DF6643: lea edx, arg_C
  loc_00DF6647: push edx
  loc_00DF6648: mov edx, [esi+00000088h]
  loc_00DF664E: lea eax, arg_C
  loc_00DF6652: mov arg_C, edi
  loc_00DF6656: push eax
  loc_00DF6657: push edx
  loc_00DF6658: push esi
  loc_00DF6659: call [ecx+0000090Ch]
  loc_00DF665F: fld real8 ptr [004013B0h]
  loc_00DF6665: fstp real4 ptr arg_10
  loc_00DF6669: fnstsw ax
  loc_00DF666B: test al, 0Dh
  loc_00DF666D: jnz 00DF6D15h
  loc_00DF6673: mov ecx, [esi]
  loc_00DF6675: lea edx, arg_18
  loc_00DF6679: mov eax, arg_C
  loc_00DF667D: push edx
  loc_00DF667E: mov arg_18, eax
  loc_00DF6682: lea eax, arg_14
  loc_00DF6686: lea edx, arg_18
  loc_00DF668A: push eax
  loc_00DF668B: push edx
  loc_00DF668C: push esi
  loc_00DF668D: call [ecx+0000097Ch]
  loc_00DF6693: mov ecx, [esi]
  loc_00DF6695: lea edx, arg_1C
  loc_00DF6699: mov eax, arg_18
  loc_00DF669D: push edx
  loc_00DF669E: push esi
  loc_00DF669F: mov arg_24, eax
  loc_00DF66A3: call [ecx+00000940h]
  loc_00DF66A9: mov eax, [esi]
  loc_00DF66AB: lea ecx, arg_C
  loc_00DF66AF: push ecx
  loc_00DF66B0: mov ecx, [esi+00000088h]
  loc_00DF66B6: lea edx, arg_C
  loc_00DF66BA: mov arg_C, edi
  loc_00DF66BE: push edx
  loc_00DF66BF: push ecx
  loc_00DF66C0: push esi
  loc_00DF66C1: call [eax+0000090Ch]
  loc_00DF66C7: fld real8 ptr [00401808h]
  loc_00DF66CD: mov edx, arg_C
  loc_00DF66D1: lea ecx, arg_18
  loc_00DF66D5: fstp real4 ptr arg_10
  loc_00DF66D9: fnstsw ax
  loc_00DF66DB: test al, 0Dh
  loc_00DF66DD: jnz 00DF6D15h
  loc_00DF66E3: mov arg_14, edx
  loc_00DF66E7: push ecx
  loc_00DF66E8: lea edx, arg_14
  loc_00DF66EC: lea ecx, arg_18
  loc_00DF66F0: push edx
  loc_00DF66F1: push ecx
  loc_00DF66F2: mov eax, [esi]
  loc_00DF66F4: push esi
  loc_00DF66F5: call [eax+0000097Ch]
  loc_00DF66FB: mov eax, [esi+000001D0h]
  loc_00DF6701: mov edx, [esi]
  loc_00DF6703: mov ecx, arg_18
  loc_00DF6707: push ecx
  loc_00DF6708: mov ecx, eax
  loc_00DF670A: sub ecx, 00000001h
  loc_00DF670D: jo 00DF6D1Ah
  loc_00DF6713: push ecx
  loc_00DF6714: mov ecx, [esi+000001D4h]
  loc_00DF671A: sub ecx, 00000002h
  loc_00DF671D: jo 00DF6D1Ah
  loc_00DF6723: sub eax, 00000001h
  loc_00DF6726: push ecx
  loc_00DF6727: jo 00DF6D1Ah
  loc_00DF672D: push eax
  loc_00DF672E: push 00000002h
  loc_00DF6730: push esi
  loc_00DF6731: call [edx+000008E4h]
  loc_00DF6737: jmp 00DF6C30h
  loc_00DF673C: mov edx, [esi]
  loc_00DF673E: lea eax, arg_C
  loc_00DF6742: lea ecx, Me
  loc_00DF6746: push eax
  loc_00DF6747: push ecx
  loc_00DF6748: push 00FFFFFFh
  loc_00DF674D: push esi
  loc_00DF674E: mov arg_18, edi
  loc_00DF6752: call [edx+0000090Ch]
  loc_00DF6758: mov edx, [esi]
  loc_00DF675A: lea eax, arg_10
  loc_00DF675E: push eax
  loc_00DF675F: mov eax, [esi+00000088h]
  loc_00DF6765: lea ecx, arg_18
  loc_00DF6769: mov arg_18, edi
  loc_00DF676D: push ecx
  loc_00DF676E: push eax
  loc_00DF676F: push esi
  loc_00DF6770: call [edx+0000090Ch]
  loc_00DF6776: fild real4 ptr [esi+000001D0h]
  loc_00DF677C: mov ecx, arg_10
  loc_00DF6780: mov edx, arg_C
  loc_00DF6784: mov ebp, [esi]
  loc_00DF6786: push 00000001h
  loc_00DF6788: fstp real8 ptr arg_34
  loc_00DF678C: fld real8 ptr arg_34
  loc_00DF6790: cmp [00E3D000h], 00000000h
  loc_00DF6797: jnz 00DF67A1h
  loc_00DF6799: fdiv st0, real8 ptr [00401338h]
  loc_00DF679F: jmp 00DF67B2h
  loc_00DF67A1: push [0040133Ch]
  loc_00DF67A7: push [00401338h]
  loc_00DF67AD: call 00402854h ; _adj_fdiv_m64
  loc_00DF67B2: push ecx
  loc_00DF67B3: push edx
  loc_00DF67B4: fnstsw ax
  loc_00DF67B6: test al, 0Dh
  loc_00DF67B8: jnz 00DF6D15h
  loc_00DF67BE: call ebx
  loc_00DF67C0: push eax
  loc_00DF67C1: mov eax, [esi+000001D4h]
  loc_00DF67C7: push eax
  loc_00DF67C8: push edi
  loc_00DF67C9: push edi
  loc_00DF67CA: push esi
  loc_00DF67CB: call arg_908
  loc_00DF67D1: mov ecx, [esi]
  loc_00DF67D3: lea edx, arg_C
  loc_00DF67D7: push edx
  loc_00DF67D8: mov edx, [esi+00000088h]
  loc_00DF67DE: lea eax, arg_C
  loc_00DF67E2: mov arg_C, edi
  loc_00DF67E6: push eax
  loc_00DF67E7: push edx
  loc_00DF67E8: push esi
  loc_00DF67E9: call [ecx+0000090Ch]
  loc_00DF67EF: mov eax, [esi]
  loc_00DF67F1: lea ecx, arg_10
  loc_00DF67F5: push ecx
  loc_00DF67F6: mov ecx, [esi+00000088h]
  loc_00DF67FC: lea edx, arg_18
  loc_00DF6800: mov arg_18, edi
  loc_00DF6804: push edx
  loc_00DF6805: push ecx
  loc_00DF6806: push esi
  loc_00DF6807: call [eax+0000090Ch]
  loc_00DF680D: mov eax, [esi+000001D0h]
  loc_00DF6813: mov ebp, [esi]
  loc_00DF6815: mov Proc_2_128_DF62E0C, eax
  loc_00DF6819: mov edx, arg_10
  loc_00DF681D: fild real4 ptr Proc_2_128_DF62E0C
  loc_00DF6821: mov ecx, arg_C
  loc_00DF6825: push 00000001h
  loc_00DF6827: push edx
  loc_00DF6828: mov edx, [esi+000001D4h]
  loc_00DF682E: fstp real8 ptr arg_38
  loc_00DF6832: fld real8 ptr arg_38
  loc_00DF6836: cmp [00E3D000h], 00000000h
  loc_00DF683D: jnz 00DF6847h
  loc_00DF683F: fdiv st0, real8 ptr [00401338h]
  loc_00DF6845: jmp 00DF6858h
  loc_00DF6847: push [0040133Ch]
  loc_00DF684D: push [00401338h]
  loc_00DF6853: call 00402854h ; _adj_fdiv_m64
  loc_00DF6858: push ecx
  loc_00DF6859: push eax
  loc_00DF685A: push edx
  loc_00DF685B: fnstsw ax
  loc_00DF685D: test al, 0Dh
  loc_00DF685F: jnz 00DF6D15h
  loc_00DF6865: call ebx
  loc_00DF6867: push eax
  loc_00DF6868: push edi
  loc_00DF6869: push esi
  loc_00DF686A: call arg_908
  loc_00DF6870: mov eax, [esi]
  loc_00DF6872: push esi
  loc_00DF6873: call [eax+00000920h]
  loc_00DF6879: mov ecx, [esi]
  loc_00DF687B: lea edx, arg_C
  loc_00DF687F: lea eax, Me
  loc_00DF6883: push edx
  loc_00DF6884: push eax
  loc_00DF6885: push 00FFFFFFh
  loc_00DF688A: push esi
  loc_00DF688B: mov arg_18, edi
  loc_00DF688F: call [ecx+0000090Ch]
  loc_00DF6895: mov eax, [esi+000001D0h]
  loc_00DF689B: mov ecx, [esi]
  loc_00DF689D: mov edx, arg_C
  loc_00DF68A1: push edx
  loc_00DF68A2: mov edx, [esi+000001D4h]
  loc_00DF68A8: sub edx, 00000002h
  loc_00DF68AB: push eax
  loc_00DF68AC: jo 00DF6D1Ah
  loc_00DF68B2: push edx
  loc_00DF68B3: push 00000001h
  loc_00DF68B5: push 00000001h
  loc_00DF68B7: push esi
  loc_00DF68B8: call [ecx+000008F0h]
  loc_00DF68BE: mov eax, [esi]
  loc_00DF68C0: lea ecx, arg_C
  loc_00DF68C4: push ecx
  loc_00DF68C5: mov ecx, [esi+00000088h]
  loc_00DF68CB: lea edx, arg_C
  loc_00DF68CF: mov arg_C, edi
  loc_00DF68D3: push edx
  loc_00DF68D4: push ecx
  loc_00DF68D5: push esi
  loc_00DF68D6: call [eax+0000090Ch]
  loc_00DF68DC: fld real8 ptr [004015A0h]
  loc_00DF68E2: mov edx, arg_C
  loc_00DF68E6: lea ecx, arg_18
  loc_00DF68EA: fstp real4 ptr arg_10
  loc_00DF68EE: fnstsw ax
  loc_00DF68F0: test al, 0Dh
  loc_00DF68F2: jnz 00DF6D15h
  loc_00DF68F8: mov arg_14, edx
  loc_00DF68FC: push ecx
  loc_00DF68FD: lea edx, arg_14
  loc_00DF6901: lea ecx, arg_18
  loc_00DF6905: mov eax, [esi]
  loc_00DF6907: push edx
  loc_00DF6908: push ecx
  loc_00DF6909: push esi
  loc_00DF690A: call [eax+0000097Ch]
  loc_00DF6910: mov ecx, [esi+000001D0h]
  loc_00DF6916: mov edx, [esi]
  loc_00DF6918: mov eax, arg_18
  loc_00DF691C: push eax
  loc_00DF691D: mov eax, [esi+000001D4h]
  loc_00DF6923: push ecx
  loc_00DF6924: push eax
  loc_00DF6925: push edi
  loc_00DF6926: push edi
  loc_00DF6927: push esi
  loc_00DF6928: call [edx+000008F0h]
  loc_00DF692E: mov ecx, [esi]
  loc_00DF6930: lea edx, arg_C
  loc_00DF6934: push edx
  loc_00DF6935: mov edx, [esi+00000088h]
  loc_00DF693B: lea eax, arg_C
  loc_00DF693F: mov arg_C, edi
  loc_00DF6943: push eax
  loc_00DF6944: push edx
  loc_00DF6945: push esi
  loc_00DF6946: call [ecx+0000090Ch]
  loc_00DF694C: fld real8 ptr [00401800h]
  loc_00DF6952: fstp real4 ptr arg_10
  loc_00DF6956: fnstsw ax
  loc_00DF6958: test al, 0Dh
  loc_00DF695A: jnz 00DF6D15h
  loc_00DF6960: mov ecx, [esi]
  loc_00DF6962: lea edx, arg_18
  loc_00DF6966: mov eax, arg_C
  loc_00DF696A: push edx
  loc_00DF696B: mov arg_18, eax
  loc_00DF696F: lea eax, arg_14
  loc_00DF6973: lea edx, arg_18
  loc_00DF6977: push eax
  loc_00DF6978: push edx
  loc_00DF6979: push esi
  loc_00DF697A: call [ecx+0000097Ch]
  loc_00DF6980: jmp 00DF6BF4h
  loc_00DF6985: mov eax, [esi]
  loc_00DF6987: push esi
  loc_00DF6988: call [eax+00000914h]
  loc_00DF698E: mov ecx, [esi+000001D0h]
  loc_00DF6994: mov edx, [esi+000001D4h]
  loc_00DF699A: push ecx
  loc_00DF699B: push edx
  loc_00DF699C: push edi
  loc_00DF699D: lea eax, [esi+000000ACh]
  loc_00DF69A3: push edi
  loc_00DF69A4: push eax
  loc_00DF69A5: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF69AA: call [00401074h] ; __vbaSetSystemError
  loc_00DF69B0: mov ecx, [esi]
  loc_00DF69B2: lea edx, arg_C
  loc_00DF69B6: lea eax, Me
  loc_00DF69BA: push edx
  loc_00DF69BB: push eax
  loc_00DF69BC: push 00FFFFFFh
  loc_00DF69C1: push esi
  loc_00DF69C2: mov arg_18, edi
  loc_00DF69C6: call [ecx+0000090Ch]
  loc_00DF69CC: mov ecx, [esi]
  loc_00DF69CE: lea edx, arg_10
  loc_00DF69D2: push edx
  loc_00DF69D3: mov edx, [esi+00000088h]
  loc_00DF69D9: lea eax, arg_18
  loc_00DF69DD: mov arg_18, edi
  loc_00DF69E1: push eax
  loc_00DF69E2: push edx
  loc_00DF69E3: push esi
  loc_00DF69E4: call [ecx+0000090Ch]
  loc_00DF69EA: fild real4 ptr [esi+000001D0h]
  loc_00DF69F0: mov eax, arg_10
  loc_00DF69F4: mov ecx, arg_C
  loc_00DF69F8: mov ebp, [esi]
  loc_00DF69FA: push 00000001h
  loc_00DF69FC: fstp real8 ptr arg_34
  loc_00DF6A00: fld real8 ptr arg_34
  loc_00DF6A04: cmp [00E3D000h], 00000000h
  loc_00DF6A0B: jnz 00DF6A15h
  loc_00DF6A0D: fdiv st0, real8 ptr [00401338h]
  loc_00DF6A13: jmp 00DF6A26h
  loc_00DF6A15: push [0040133Ch]
  loc_00DF6A1B: push [00401338h]
  loc_00DF6A21: call 00402854h ; _adj_fdiv_m64
  loc_00DF6A26: push eax
  loc_00DF6A27: push ecx
  loc_00DF6A28: fnstsw ax
  loc_00DF6A2A: test al, 0Dh
  loc_00DF6A2C: jnz 00DF6D15h
  loc_00DF6A32: call ebx
  loc_00DF6A34: mov edx, [esi+000001D4h]
  loc_00DF6A3A: push eax
  loc_00DF6A3B: push edx
  loc_00DF6A3C: push edi
  loc_00DF6A3D: push edi
  loc_00DF6A3E: push esi
  loc_00DF6A3F: call arg_908
  loc_00DF6A45: mov eax, [esi]
  loc_00DF6A47: lea ecx, arg_C
  loc_00DF6A4B: push ecx
  loc_00DF6A4C: mov ecx, [esi+00000088h]
  loc_00DF6A52: lea edx, arg_C
  loc_00DF6A56: mov arg_C, edi
  loc_00DF6A5A: push edx
  loc_00DF6A5B: push ecx
  loc_00DF6A5C: push esi
  loc_00DF6A5D: call [eax+0000090Ch]
  loc_00DF6A63: mov edx, [esi]
  loc_00DF6A65: lea eax, arg_10
  loc_00DF6A69: push eax
  loc_00DF6A6A: mov eax, [esi+00000088h]
  loc_00DF6A70: lea ecx, arg_18
  loc_00DF6A74: mov arg_18, edi
  loc_00DF6A78: push ecx
  loc_00DF6A79: push eax
  loc_00DF6A7A: push esi
  loc_00DF6A7B: call [edx+0000090Ch]
  loc_00DF6A81: mov eax, [esi+000001D0h]
  loc_00DF6A87: mov ebp, [esi]
  loc_00DF6A89: mov ecx, arg_10
  loc_00DF6A8D: mov edx, arg_C
  loc_00DF6A91: push 00000001h
  loc_00DF6A93: push ecx
  loc_00DF6A94: push edx
  loc_00DF6A95: mov arg_58, eax
  loc_00DF6A99: push eax
  loc_00DF6A9A: fild real4 ptr arg_5C
  loc_00DF6A9E: mov eax, [esi+000001D4h]
  loc_00DF6AA4: push eax
  loc_00DF6AA5: fstp real8 ptr Proc_2_128_DF62E04
  loc_00DF6AA9: fld real8 ptr Proc_2_128_DF62E04
  loc_00DF6AAD: cmp [00E3D000h], 00000000h
  loc_00DF6AB4: jnz 00DF6ABEh
  loc_00DF6AB6: fdiv st0, real8 ptr [00401338h]
  loc_00DF6ABC: jmp 00DF6ACFh
  loc_00DF6ABE: push [0040133Ch]
  loc_00DF6AC4: push [00401338h]
  loc_00DF6ACA: call 00402854h ; _adj_fdiv_m64
  loc_00DF6ACF: fnstsw ax
  loc_00DF6AD1: test al, 0Dh
  loc_00DF6AD3: jnz 00DF6D15h
  loc_00DF6AD9: call ebx
  loc_00DF6ADB: push eax
  loc_00DF6ADC: push edi
  loc_00DF6ADD: push esi
  loc_00DF6ADE: call arg_908
  loc_00DF6AE4: mov ecx, [esi]
  loc_00DF6AE6: push esi
  loc_00DF6AE7: call [ecx+00000920h]
  loc_00DF6AED: mov edx, [esi]
  loc_00DF6AEF: lea eax, arg_C
  loc_00DF6AF3: lea ecx, Me
  loc_00DF6AF7: push eax
  loc_00DF6AF8: push ecx
  loc_00DF6AF9: push 00FFFFFFh
  loc_00DF6AFE: push esi
  loc_00DF6AFF: mov arg_18, edi
  loc_00DF6B03: call [edx+0000090Ch]
  loc_00DF6B09: mov ecx, [esi+000001D0h]
  loc_00DF6B0F: mov edx, [esi]
  loc_00DF6B11: mov eax, arg_C
  loc_00DF6B15: push eax
  loc_00DF6B16: mov eax, [esi+000001D4h]
  loc_00DF6B1C: sub eax, 00000002h
  loc_00DF6B1F: push ecx
  loc_00DF6B20: jo 00DF6D1Ah
  loc_00DF6B26: push eax
  loc_00DF6B27: push 00000001h
  loc_00DF6B29: push 00000001h
  loc_00DF6B2B: push esi
  loc_00DF6B2C: call [edx+000008F0h]
  loc_00DF6B32: mov ecx, [esi]
  loc_00DF6B34: lea edx, arg_C
  loc_00DF6B38: push edx
  loc_00DF6B39: mov edx, [esi+00000088h]
  loc_00DF6B3F: lea eax, arg_C
  loc_00DF6B43: mov arg_C, edi
  loc_00DF6B47: push eax
  loc_00DF6B48: push edx
  loc_00DF6B49: push esi
  loc_00DF6B4A: call [ecx+0000090Ch]
  loc_00DF6B50: fld real8 ptr [004015A0h]
  loc_00DF6B56: fstp real4 ptr arg_10
  loc_00DF6B5A: fnstsw ax
  loc_00DF6B5C: test al, 0Dh
  loc_00DF6B5E: jnz 00DF6D15h
  loc_00DF6B64: mov ecx, [esi]
  loc_00DF6B66: lea edx, arg_18
  loc_00DF6B6A: mov eax, arg_C
  loc_00DF6B6E: push edx
  loc_00DF6B6F: mov arg_18, eax
  loc_00DF6B73: lea eax, arg_14
  loc_00DF6B77: lea edx, arg_18
  loc_00DF6B7B: push eax
  loc_00DF6B7C: push edx
  loc_00DF6B7D: push esi
  loc_00DF6B7E: call [ecx+0000097Ch]
  loc_00DF6B84: mov edx, [esi+000001D0h]
  loc_00DF6B8A: mov eax, [esi]
  loc_00DF6B8C: mov ecx, arg_18
  loc_00DF6B90: push ecx
  loc_00DF6B91: mov ecx, [esi+000001D4h]
  loc_00DF6B97: push edx
  loc_00DF6B98: push ecx
  loc_00DF6B99: push edi
  loc_00DF6B9A: push edi
  loc_00DF6B9B: push esi
  loc_00DF6B9C: call [eax+000008F0h]
  loc_00DF6BA2: mov edx, [esi]
  loc_00DF6BA4: lea eax, arg_C
  loc_00DF6BA8: lea ecx, Me
  loc_00DF6BAC: push eax
  loc_00DF6BAD: mov eax, [esi+00000088h]
  loc_00DF6BB3: mov arg_C, edi
  loc_00DF6BB7: push ecx
  loc_00DF6BB8: push eax
  loc_00DF6BB9: push esi
  loc_00DF6BBA: call [edx+0000090Ch]
  loc_00DF6BC0: fld real8 ptr [00401800h]
  loc_00DF6BC6: mov ecx, arg_C
  loc_00DF6BCA: mov edx, [esi]
  loc_00DF6BCC: fstp real4 ptr arg_10
  loc_00DF6BD0: fnstsw ax
  loc_00DF6BD2: test al, 0Dh
  loc_00DF6BD4: jnz 00DF6D15h
  loc_00DF6BDA: lea eax, arg_18
  loc_00DF6BDE: mov arg_14, ecx
  loc_00DF6BE2: push eax
  loc_00DF6BE3: lea ecx, arg_14
  loc_00DF6BE7: lea eax, arg_18
  loc_00DF6BEB: push ecx
  loc_00DF6BEC: push eax
  loc_00DF6BED: push esi
  loc_00DF6BEE: call [edx+0000097Ch]
  loc_00DF6BF4: mov edx, arg_18
  loc_00DF6BF8: mov eax, [esi+000001D0h]
  loc_00DF6BFE: push edx
  loc_00DF6BFF: mov edx, eax
  loc_00DF6C01: sub edx, 00000001h
  loc_00DF6C04: mov ecx, [esi]
  loc_00DF6C06: jo 00DF6D1Ah
  loc_00DF6C0C: push edx
  loc_00DF6C0D: mov edx, [esi+000001D4h]
  loc_00DF6C13: sub edx, 00000002h
  loc_00DF6C16: jo 00DF6D1Ah
  loc_00DF6C1C: sub eax, 00000001h
  loc_00DF6C1F: push edx
  loc_00DF6C20: jo 00DF6D1Ah
  loc_00DF6C26: push eax
  loc_00DF6C27: push 00000002h
  loc_00DF6C29: push esi
  loc_00DF6C2A: call [ecx+000008E4h]
  loc_00DF6C30: mov eax, [esi]
  loc_00DF6C32: lea ecx, arg_C
  loc_00DF6C36: push ecx
  loc_00DF6C37: mov ecx, [esi+00000088h]
  loc_00DF6C3D: lea edx, arg_C
  loc_00DF6C41: mov arg_C, edi
  loc_00DF6C45: push edx
  loc_00DF6C46: push ecx
  loc_00DF6C47: push esi
  loc_00DF6C48: call [eax+0000090Ch]
  loc_00DF6C4E: mov eax, [esi]
  loc_00DF6C50: lea ecx, arg_18
  loc_00DF6C54: mov edx, arg_C
  loc_00DF6C58: push ecx
  loc_00DF6C59: mov arg_18, edx
  loc_00DF6C5D: lea edx, arg_14
  loc_00DF6C61: lea ecx, arg_18
  loc_00DF6C65: push edx
  loc_00DF6C66: push ecx
  loc_00DF6C67: push esi
  loc_00DF6C68: mov arg_20, 3D4CCCCDh
  loc_00DF6C70: call [eax+0000097Ch]
  loc_00DF6C76: mov eax, [esi]
  loc_00DF6C78: lea ecx, arg_1C
  loc_00DF6C7C: mov edx, arg_18
  loc_00DF6C80: push ecx
  loc_00DF6C81: push esi
  loc_00DF6C82: mov arg_24, edx
  loc_00DF6C86: call [eax+00000940h]
  loc_00DF6C8C: cmp [esi+00000070h], di
  loc_00DF6C90: jz 00DF6D09h
  loc_00DF6C92: cmp [esi+0000006Eh], di
  loc_00DF6C96: jz 00DF6D09h
  loc_00DF6C98: cmp [esi+00000050h], di
  loc_00DF6C9C: jnz 00DF6CA4h
  loc_00DF6C9E: cmp [esi+00000058h], di
  loc_00DF6CA2: jz 00DF6D09h
  loc_00DF6CA4: mov edx, [esi+000001D0h]
  loc_00DF6CAA: mov eax, [esi+000001D4h]
  loc_00DF6CB0: sub edx, 00000003h
  loc_00DF6CB3: lea ecx, arg_38
  loc_00DF6CB7: jo 00DF6D1Ah
  loc_00DF6CB9: sub eax, 00000003h
  loc_00DF6CBC: push edx
  loc_00DF6CBD: jo 00DF6D1Ah
  loc_00DF6CBF: push eax
  loc_00DF6CC0: push 00000003h
  loc_00DF6CC2: push 00000003h
  loc_00DF6CC4: push ecx
  loc_00DF6CC5: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF6CCA: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00DF6CD0: call ebx
  loc_00DF6CD2: mov edx, [esi]
  loc_00DF6CD4: lea eax, Me
  loc_00DF6CD8: push eax
  loc_00DF6CD9: push esi
  loc_00DF6CDA: call [edx+000000D8h]
  loc_00DF6CE0: cmp eax, edi
  loc_00DF6CE2: fnclex
  loc_00DF6CE4: jge 00DF6CF8h
  loc_00DF6CE6: push 000000D8h
  loc_00DF6CEB: push 006DCB00h
  loc_00DF6CF0: push esi
  loc_00DF6CF1: push eax
  loc_00DF6CF2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF6CF8: mov edx, Me
  loc_00DF6CFC: lea ecx, arg_38
  loc_00DF6D00: push ecx
  loc_00DF6D01: push edx
  loc_00DF6D02: call 006DDA18h ; DrawFocusRect(%x1v, %x2v)
  loc_00DF6D07: call ebx
  loc_00DF6D09: pop edi
  loc_00DF6D0A: pop esi
  loc_00DF6D0B: pop ebp
  loc_00DF6D0C: xor eax, eax
  loc_00DF6D0E: pop ebx
  loc_00DF6D0F: add esp, 00000040h
  loc_00DF6D12: retn 0008h
End Sub

Private Sub Proc_2_129_DF6D20(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28, arg_2C, arg_30, arg_34) 'DF6D20
  loc_00DF6D20: sub esp, 00000030h
  loc_00DF6D23: xor eax, eax
  loc_00DF6D25: push ebx
  loc_00DF6D26: push esi
  loc_00DF6D27: mov esi, arg_34
  loc_00DF6D2B: mov arg_20, eax
  loc_00DF6D2F: lea edx, [esp+00000008h]
  loc_00DF6D33: mov ecx, [esi]
  loc_00DF6D35: mov arg_24, eax
  loc_00DF6D39: push edi
  loc_00DF6D3A: mov arg_2C, eax
  loc_00DF6D3E: xor edi, edi
  loc_00DF6D40: push edx
  loc_00DF6D41: push esi
  loc_00DF6D42: mov arg_38, eax
  loc_00DF6D46: mov arg_14, edi
  loc_00DF6D4A: mov arg_C, edi
  loc_00DF6D4E: mov arg_10, edi
  loc_00DF6D52: mov arg_18, edi
  loc_00DF6D56: mov arg_24, edi
  loc_00DF6D5A: mov arg_1C, edi
  loc_00DF6D5E: mov arg_20, edi
  loc_00DF6D62: mov arg_28, edi
  loc_00DF6D66: call [ecx+00000108h]
  loc_00DF6D6C: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF6D72: cmp eax, edi
  loc_00DF6D74: fnclex
  loc_00DF6D76: jge 00DF6D86h
  loc_00DF6D78: push 00000108h
  loc_00DF6D7D: push 006DCB00h
  loc_00DF6D82: push esi
  loc_00DF6D83: push eax
  loc_00DF6D84: call ebx
  loc_00DF6D86: fld real4 ptr Proc_2_129_DF6D20
  loc_00DF6D8A: push ebp
  loc_00DF6D8B: mov ebp, [00401208h] ; __vbaFpI4
  loc_00DF6D91: call ebp
  loc_00DF6D93: lea ecx, Me
  loc_00DF6D97: mov [esi+000001D0h], eax
  loc_00DF6D9D: mov eax, [esi]
  loc_00DF6D9F: push ecx
  loc_00DF6DA0: push esi
  loc_00DF6DA1: call [eax+00000100h]
  loc_00DF6DA7: cmp eax, edi
  loc_00DF6DA9: fnclex
  loc_00DF6DAB: jge 00DF6DBBh
  loc_00DF6DAD: push 00000100h
  loc_00DF6DB2: push 006DCB00h
  loc_00DF6DB7: push esi
  loc_00DF6DB8: push eax
  loc_00DF6DB9: call ebx
  loc_00DF6DBB: fld real4 ptr Me
  loc_00DF6DBF: call ebp
  loc_00DF6DC1: mov edx, [esi]
  loc_00DF6DC3: mov [esi+000001D4h], eax
  loc_00DF6DC9: lea eax, arg_C
  loc_00DF6DCD: lea ecx, Me
  loc_00DF6DD1: push eax
  loc_00DF6DD2: mov eax, [esi+00000088h]
  loc_00DF6DD8: push ecx
  loc_00DF6DD9: push eax
  loc_00DF6DDA: push esi
  loc_00DF6DDB: mov arg_18, edi
  loc_00DF6DDF: call [edx+0000090Ch]
  loc_00DF6DE5: cmp [esi+0000007Ch], di
  loc_00DF6DE9: pop ebp
  loc_00DF6DEA: mov ecx, Me
  loc_00DF6DEE: mov arg_C, ecx
  loc_00DF6DF2: jnz 00DF6FB4h
  loc_00DF6DF8: mov edx, [esi+000001D0h]
  loc_00DF6DFE: mov eax, [esi+000001D4h]
  loc_00DF6E04: push edx
  loc_00DF6E05: push eax
  loc_00DF6E06: push edi
  loc_00DF6E07: lea ecx, arg_30
  loc_00DF6E0B: push edi
  loc_00DF6E0C: push ecx
  loc_00DF6E0D: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF6E12: call [00401074h] ; __vbaSetSystemError
  loc_00DF6E18: mov ecx, arg_C
  loc_00DF6E1C: mov edx, [esi]
  loc_00DF6E1E: lea eax, arg_24
  loc_00DF6E22: push eax
  loc_00DF6E23: push ecx
  loc_00DF6E24: push esi
  loc_00DF6E25: call [edx+00000974h]
  loc_00DF6E2B: mov edx, [esi]
  loc_00DF6E2D: lea eax, Me
  loc_00DF6E31: lea ecx, Proc_2_129_DF6D20
  loc_00DF6E35: push eax
  loc_00DF6E36: push ecx
  loc_00DF6E37: push 00FFFFFFh
  loc_00DF6E3C: push esi
  loc_00DF6E3D: mov arg_14, edi
  loc_00DF6E41: call [edx+0000090Ch]
  loc_00DF6E47: mov edx, [esi]
  loc_00DF6E49: lea eax, arg_10
  loc_00DF6E4D: mov ecx, Me
  loc_00DF6E51: push eax
  loc_00DF6E52: mov eax, arg_10
  loc_00DF6E56: push ecx
  loc_00DF6E57: push eax
  loc_00DF6E58: push esi
  loc_00DF6E59: call [edx+000008ECh]
  loc_00DF6E5F: mov edx, [esi]
  loc_00DF6E61: lea eax, arg_18
  loc_00DF6E65: mov ecx, arg_10
  loc_00DF6E69: push eax
  loc_00DF6E6A: mov arg_20, ecx
  loc_00DF6E6E: lea ecx, arg_18
  loc_00DF6E72: lea eax, arg_20
  loc_00DF6E76: push ecx
  loc_00DF6E77: push eax
  loc_00DF6E78: push esi
  loc_00DF6E79: mov arg_24, 3D4CCCCDh
  loc_00DF6E81: call [edx+0000097Ch]
  loc_00DF6E87: mov ecx, [esi]
  loc_00DF6E89: push 00000001h
  loc_00DF6E8B: mov edx, arg_10
  loc_00DF6E8F: mov eax, arg_1C
  loc_00DF6E93: push edx
  loc_00DF6E94: mov edx, [esi+000001D4h]
  loc_00DF6E9A: push eax
  loc_00DF6E9B: push 00000005h
  loc_00DF6E9D: push edx
  loc_00DF6E9E: push edi
  loc_00DF6E9F: push edi
  loc_00DF6EA0: push esi
  loc_00DF6EA1: call [ecx+00000908h]
  loc_00DF6EA7: fld real8 ptr [00401598h]
  loc_00DF6EAD: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF6EB1: fnstsw ax
  loc_00DF6EB3: test al, 0Dh
  loc_00DF6EB5: jnz 00DF7576h
  loc_00DF6EBB: lea ecx, Me
  loc_00DF6EBF: lea edx, Proc_2_129_DF6D20
  loc_00DF6EC3: push ecx
  loc_00DF6EC4: lea ecx, arg_10
  loc_00DF6EC8: mov eax, [esi]
  loc_00DF6ECA: push edx
  loc_00DF6ECB: push ecx
  loc_00DF6ECC: push esi
  loc_00DF6ECD: call [eax+0000097Ch]
  loc_00DF6ED3: mov edx, [esi]
  loc_00DF6ED5: lea eax, arg_1C
  loc_00DF6ED9: lea ecx, arg_10
  loc_00DF6EDD: push eax
  loc_00DF6EDE: push ecx
  loc_00DF6EDF: push 00FFFFFFh
  loc_00DF6EE4: mov arg_1C, edi
  loc_00DF6EE8: push esi
  loc_00DF6EE9: call [edx+0000090Ch]
  loc_00DF6EEF: mov edx, [esi]
  loc_00DF6EF1: lea eax, arg_18
  loc_00DF6EF5: push eax
  loc_00DF6EF6: lea ecx, arg_18
  loc_00DF6EFA: lea eax, arg_10
  loc_00DF6EFE: push ecx
  loc_00DF6EFF: push eax
  loc_00DF6F00: push esi
  loc_00DF6F01: mov arg_24, 3DA3D70Ah
  loc_00DF6F09: call [edx+0000097Ch]
  loc_00DF6F0F: mov ecx, [esi]
  loc_00DF6F11: lea edx, arg_20
  loc_00DF6F15: mov eax, arg_18
  loc_00DF6F19: push edx
  loc_00DF6F1A: mov edx, arg_20
  loc_00DF6F1E: push eax
  loc_00DF6F1F: push edx
  loc_00DF6F20: push esi
  loc_00DF6F21: call [ecx+000008ECh]
  loc_00DF6F27: push 00000001h
  loc_00DF6F29: mov eax, [esi]
  loc_00DF6F2B: mov ecx, arg_24
  loc_00DF6F2F: mov edx, arg_C
  loc_00DF6F33: push ecx
  loc_00DF6F34: mov ecx, [esi+000001D0h]
  loc_00DF6F3A: sub ecx, 00000001h
  loc_00DF6F3D: push edx
  loc_00DF6F3E: mov edx, [esi+000001D4h]
  loc_00DF6F44: jo 00DF757Bh
  loc_00DF6F4A: push ecx
  loc_00DF6F4B: push edx
  loc_00DF6F4C: push 00000006h
  loc_00DF6F4E: push edi
  loc_00DF6F4F: push esi
  loc_00DF6F50: call [eax+00000908h]
  loc_00DF6F56: mov eax, [esi]
  loc_00DF6F58: push esi
  loc_00DF6F59: call [eax+00000920h]
  loc_00DF6F5F: fld real8 ptr [004015A0h]
  loc_00DF6F65: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF6F69: fnstsw ax
  loc_00DF6F6B: test al, 0Dh
  loc_00DF6F6D: jnz 00DF7576h
  loc_00DF6F73: mov ecx, [esi]
  loc_00DF6F75: lea edx, Me
  loc_00DF6F79: push edx
  loc_00DF6F7A: lea eax, Me
  loc_00DF6F7E: lea edx, arg_10
  loc_00DF6F82: push eax
  loc_00DF6F83: push edx
  loc_00DF6F84: push esi
  loc_00DF6F85: call [ecx+0000097Ch]
  loc_00DF6F8B: mov edx, [esi+000001D0h]
  loc_00DF6F91: mov eax, [esi]
  loc_00DF6F93: mov ecx, Me
  loc_00DF6F97: push ecx
  loc_00DF6F98: mov ecx, [esi+000001D4h]
  loc_00DF6F9E: push edx
  loc_00DF6F9F: push ecx
  loc_00DF6FA0: push edi
  loc_00DF6FA1: push edi
  loc_00DF6FA2: push esi
  loc_00DF6FA3: call [eax+000008F0h]
  loc_00DF6FA9: fld real8 ptr [00401810h]
  loc_00DF6FAF: jmp 00DF7533h
  loc_00DF6FB4: mov eax, arg_3C
  loc_00DF6FB8: sub eax, edi
  loc_00DF6FBA: jz 00DF734Eh
  loc_00DF6FC0: dec eax
  loc_00DF6FC1: jz 00DF71B0h
  loc_00DF6FC7: dec eax
  loc_00DF6FC8: jnz 00DF752Dh
  loc_00DF6FCE: mov ecx, [esi+000001D0h]
  loc_00DF6FD4: mov edx, [esi+000001D4h]
  loc_00DF6FDA: push ecx
  loc_00DF6FDB: push edx
  loc_00DF6FDC: push edi
  loc_00DF6FDD: lea eax, arg_30
  loc_00DF6FE1: push edi
  loc_00DF6FE2: push eax
  loc_00DF6FE3: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF6FE8: call [00401074h] ; __vbaSetSystemError
  loc_00DF6FEE: fld real8 ptr [00401830h]
  loc_00DF6FF4: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF6FF8: fnstsw ax
  loc_00DF6FFA: test al, 0Dh
  loc_00DF6FFC: jnz 00DF7576h
  loc_00DF7002: mov ecx, [esi]
  loc_00DF7004: lea edx, Me
  loc_00DF7008: push edx
  loc_00DF7009: lea eax, Me
  loc_00DF700D: lea edx, arg_10
  loc_00DF7011: push eax
  loc_00DF7012: push edx
  loc_00DF7013: push esi
  loc_00DF7014: call [ecx+0000097Ch]
  loc_00DF701A: mov eax, [esi]
  loc_00DF701C: lea ecx, arg_24
  loc_00DF7020: mov edx, Me
  loc_00DF7024: push ecx
  loc_00DF7025: push edx
  loc_00DF7026: push esi
  loc_00DF7027: call [eax+00000974h]
  loc_00DF702D: mov eax, [esi]
  loc_00DF702F: lea ecx, Me
  loc_00DF7033: lea edx, Proc_2_129_DF6D20
  loc_00DF7037: push ecx
  loc_00DF7038: push edx
  loc_00DF7039: push 00FFFFFFh
  loc_00DF703E: push esi
  loc_00DF703F: mov arg_14, edi
  loc_00DF7043: call [eax+0000090Ch]
  loc_00DF7049: mov eax, [esi]
  loc_00DF704B: lea ecx, arg_10
  loc_00DF704F: mov edx, Me
  loc_00DF7053: push ecx
  loc_00DF7054: mov ecx, arg_10
  loc_00DF7058: push edx
  loc_00DF7059: push ecx
  loc_00DF705A: push esi
  loc_00DF705B: call [eax+000008ECh]
  loc_00DF7061: mov eax, [esi]
  loc_00DF7063: lea ecx, arg_18
  loc_00DF7067: mov edx, arg_10
  loc_00DF706B: push ecx
  loc_00DF706C: mov arg_20, edx
  loc_00DF7070: lea edx, arg_18
  loc_00DF7074: lea ecx, arg_20
  loc_00DF7078: push edx
  loc_00DF7079: push ecx
  loc_00DF707A: push esi
  loc_00DF707B: mov arg_24, 3DCCCCCDh
  loc_00DF7083: call [eax+0000097Ch]
  loc_00DF7089: mov edx, [esi]
  loc_00DF708B: push 00000001h
  loc_00DF708D: mov eax, arg_10
  loc_00DF7091: mov ecx, arg_1C
  loc_00DF7095: push eax
  loc_00DF7096: mov eax, [esi+000001D4h]
  loc_00DF709C: push ecx
  loc_00DF709D: push 00000005h
  loc_00DF709F: push eax
  loc_00DF70A0: push edi
  loc_00DF70A1: push edi
  loc_00DF70A2: push esi
  loc_00DF70A3: call [edx+00000908h]
  loc_00DF70A9: fld real8 ptr [004017E8h]
  loc_00DF70AF: mov ecx, [esi]
  loc_00DF70B1: lea edx, Me
  loc_00DF70B5: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF70B9: fnstsw ax
  loc_00DF70BB: test al, 0Dh
  loc_00DF70BD: jnz 00DF7576h
  loc_00DF70C3: push edx
  loc_00DF70C4: lea eax, Me
  loc_00DF70C8: lea edx, arg_10
  loc_00DF70CC: push eax
  loc_00DF70CD: push edx
  loc_00DF70CE: push esi
  loc_00DF70CF: call [ecx+0000097Ch]
  loc_00DF70D5: mov eax, [esi]
  loc_00DF70D7: lea ecx, arg_1C
  loc_00DF70DB: lea edx, arg_10
  loc_00DF70DF: push ecx
  loc_00DF70E0: push edx
  loc_00DF70E1: push 00FFFFFFh
  loc_00DF70E6: push esi
  loc_00DF70E7: mov arg_20, edi
  loc_00DF70EB: call [eax+0000090Ch]
  loc_00DF70F1: mov eax, [esi]
  loc_00DF70F3: lea ecx, arg_18
  loc_00DF70F7: push ecx
  loc_00DF70F8: lea edx, arg_18
  loc_00DF70FC: lea ecx, arg_10
  loc_00DF7100: push edx
  loc_00DF7101: push ecx
  loc_00DF7102: push esi
  loc_00DF7103: mov arg_24, 3D4CCCCDh
  loc_00DF710B: call [eax+0000097Ch]
  loc_00DF7111: mov edx, [esi]
  loc_00DF7113: lea eax, arg_20
  loc_00DF7117: mov ecx, arg_18
  loc_00DF711B: push eax
  loc_00DF711C: mov eax, arg_20
  loc_00DF7120: push ecx
  loc_00DF7121: push eax
  loc_00DF7122: push esi
  loc_00DF7123: call [edx+000008ECh]
  loc_00DF7129: push 00000001h
  loc_00DF712B: mov ecx, [esi]
  loc_00DF712D: mov edx, arg_24
  loc_00DF7131: mov eax, arg_C
  loc_00DF7135: push edx
  loc_00DF7136: mov edx, [esi+000001D0h]
  loc_00DF713C: sub edx, 00000001h
  loc_00DF713F: push eax
  loc_00DF7140: mov eax, [esi+000001D4h]
  loc_00DF7146: jo 00DF757Bh
  loc_00DF714C: push edx
  loc_00DF714D: push eax
  loc_00DF714E: push 00000006h
  loc_00DF7150: push edi
  loc_00DF7151: push esi
  loc_00DF7152: call [ecx+00000908h]
  loc_00DF7158: mov ecx, [esi]
  loc_00DF715A: push esi
  loc_00DF715B: call [ecx+00000920h]
  loc_00DF7161: fld real8 ptr [00401828h]
  loc_00DF7167: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF716B: fnstsw ax
  loc_00DF716D: test al, 0Dh
  loc_00DF716F: jnz 00DF7576h
  loc_00DF7175: mov edx, [esi]
  loc_00DF7177: lea ecx, Proc_2_129_DF6D20
  loc_00DF717B: lea eax, Me
  loc_00DF717F: push eax
  loc_00DF7180: lea eax, arg_10
  loc_00DF7184: push ecx
  loc_00DF7185: push eax
  loc_00DF7186: push esi
  loc_00DF7187: call [edx+0000097Ch]
  loc_00DF718D: mov eax, [esi+000001D0h]
  loc_00DF7193: mov ecx, [esi]
  loc_00DF7195: mov edx, Me
  loc_00DF7199: push edx
  loc_00DF719A: mov edx, [esi+000001D4h]
  loc_00DF71A0: push eax
  loc_00DF71A1: push edx
  loc_00DF71A2: push edi
  loc_00DF71A3: push edi
  loc_00DF71A4: push esi
  loc_00DF71A5: call [ecx+000008F0h]
  loc_00DF71AB: jmp 00DF752Dh
  loc_00DF71B0: mov eax, [esi+000001D0h]
  loc_00DF71B6: mov ecx, [esi+000001D4h]
  loc_00DF71BC: push eax
  loc_00DF71BD: push ecx
  loc_00DF71BE: push edi
  loc_00DF71BF: lea edx, arg_30
  loc_00DF71C3: push edi
  loc_00DF71C4: push edx
  loc_00DF71C5: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF71CA: call [00401074h] ; __vbaSetSystemError
  loc_00DF71D0: fld real8 ptr [00401830h]
  loc_00DF71D6: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF71DA: fnstsw ax
  loc_00DF71DC: test al, 0Dh
  loc_00DF71DE: jnz 00DF7576h
  loc_00DF71E4: lea ecx, Me
  loc_00DF71E8: lea edx, Proc_2_129_DF6D20
  loc_00DF71EC: push ecx
  loc_00DF71ED: lea ecx, arg_10
  loc_00DF71F1: mov eax, [esi]
  loc_00DF71F3: push edx
  loc_00DF71F4: push ecx
  loc_00DF71F5: push esi
  loc_00DF71F6: call [eax+0000097Ch]
  loc_00DF71FC: mov edx, [esi]
  loc_00DF71FE: lea eax, arg_24
  loc_00DF7202: mov ecx, Me
  loc_00DF7206: push eax
  loc_00DF7207: push ecx
  loc_00DF7208: push esi
  loc_00DF7209: call [edx+00000974h]
  loc_00DF720F: mov edx, [esi]
  loc_00DF7211: lea eax, Me
  loc_00DF7215: lea ecx, Proc_2_129_DF6D20
  loc_00DF7219: push eax
  loc_00DF721A: push ecx
  loc_00DF721B: push 00FFFFFFh
  loc_00DF7220: push esi
  loc_00DF7221: mov arg_14, edi
  loc_00DF7225: call [edx+0000090Ch]
  loc_00DF722B: mov edx, [esi]
  loc_00DF722D: lea eax, arg_10
  loc_00DF7231: mov ecx, Me
  loc_00DF7235: push eax
  loc_00DF7236: mov eax, arg_10
  loc_00DF723A: push ecx
  loc_00DF723B: push eax
  loc_00DF723C: push esi
  loc_00DF723D: call [edx+000008ECh]
  loc_00DF7243: mov edx, [esi]
  loc_00DF7245: lea eax, arg_18
  loc_00DF7249: mov ecx, arg_10
  loc_00DF724D: push eax
  loc_00DF724E: mov arg_20, ecx
  loc_00DF7252: lea ecx, arg_18
  loc_00DF7256: lea eax, arg_20
  loc_00DF725A: push ecx
  loc_00DF725B: push eax
  loc_00DF725C: push esi
  loc_00DF725D: mov arg_24, 3E19999Ah
  loc_00DF7265: call [edx+0000097Ch]
  loc_00DF726B: mov ecx, [esi]
  loc_00DF726D: push 00000001h
  loc_00DF726F: mov edx, arg_10
  loc_00DF7273: mov eax, arg_1C
  loc_00DF7277: push edx
  loc_00DF7278: mov edx, [esi+000001D4h]
  loc_00DF727E: push eax
  loc_00DF727F: push 00000005h
  loc_00DF7281: push edx
  loc_00DF7282: push edi
  loc_00DF7283: push edi
  loc_00DF7284: push esi
  loc_00DF7285: call [ecx+00000908h]
  loc_00DF728B: fld real8 ptr [00401590h]
  loc_00DF7291: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF7295: fnstsw ax
  loc_00DF7297: test al, 0Dh
  loc_00DF7299: jnz 00DF7576h
  loc_00DF729F: lea ecx, Me
  loc_00DF72A3: lea edx, Proc_2_129_DF6D20
  loc_00DF72A7: mov eax, [esi]
  loc_00DF72A9: push ecx
  loc_00DF72AA: lea ecx, arg_10
  loc_00DF72AE: push edx
  loc_00DF72AF: push ecx
  loc_00DF72B0: push esi
  loc_00DF72B1: call [eax+0000097Ch]
  loc_00DF72B7: mov edx, [esi]
  loc_00DF72B9: lea eax, arg_1C
  loc_00DF72BD: lea ecx, arg_10
  loc_00DF72C1: push eax
  loc_00DF72C2: push ecx
  loc_00DF72C3: push 00FFFFFFh
  loc_00DF72C8: push esi
  loc_00DF72C9: mov arg_20, edi
  loc_00DF72CD: call [edx+0000090Ch]
  loc_00DF72D3: mov edx, [esi]
  loc_00DF72D5: lea eax, arg_18
  loc_00DF72D9: push eax
  loc_00DF72DA: lea ecx, arg_18
  loc_00DF72DE: lea eax, arg_10
  loc_00DF72E2: push ecx
  loc_00DF72E3: push eax
  loc_00DF72E4: push esi
  loc_00DF72E5: mov arg_24, 3E4CCCCDh
  loc_00DF72ED: call [edx+0000097Ch]
  loc_00DF72F3: mov ecx, [esi]
  loc_00DF72F5: lea edx, arg_20
  loc_00DF72F9: mov eax, arg_18
  loc_00DF72FD: push edx
  loc_00DF72FE: mov edx, arg_20
  loc_00DF7302: push eax
  loc_00DF7303: push edx
  loc_00DF7304: push esi
  loc_00DF7305: call [ecx+000008ECh]
  loc_00DF730B: push 00000001h
  loc_00DF730D: mov eax, [esi]
  loc_00DF730F: mov ecx, arg_24
  loc_00DF7313: mov edx, arg_C
  loc_00DF7317: push ecx
  loc_00DF7318: mov ecx, [esi+000001D0h]
  loc_00DF731E: sub ecx, 00000001h
  loc_00DF7321: push edx
  loc_00DF7322: mov edx, [esi+000001D4h]
  loc_00DF7328: jo 00DF757Bh
  loc_00DF732E: push ecx
  loc_00DF732F: push edx
  loc_00DF7330: push 00000006h
  loc_00DF7332: push edi
  loc_00DF7333: push esi
  loc_00DF7334: call [eax+00000908h]
  loc_00DF733A: mov eax, [esi]
  loc_00DF733C: push esi
  loc_00DF733D: call [eax+00000920h]
  loc_00DF7343: fld real8 ptr [00401820h]
  loc_00DF7349: jmp 00DF74EDh
  loc_00DF734E: mov edx, [esi]
  loc_00DF7350: push esi
  loc_00DF7351: call [edx+00000914h]
  loc_00DF7357: mov eax, [esi+000001D0h]
  loc_00DF735D: mov ecx, [esi+000001D4h]
  loc_00DF7363: push eax
  loc_00DF7364: push ecx
  loc_00DF7365: push edi
  loc_00DF7366: lea edx, arg_30
  loc_00DF736A: push edi
  loc_00DF736B: push edx
  loc_00DF736C: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF7371: call [00401074h] ; __vbaSetSystemError
  loc_00DF7377: fld real8 ptr [00401830h]
  loc_00DF737D: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF7381: fnstsw ax
  loc_00DF7383: test al, 0Dh
  loc_00DF7385: jnz 00DF7576h
  loc_00DF738B: lea ecx, Me
  loc_00DF738F: lea edx, Proc_2_129_DF6D20
  loc_00DF7393: push ecx
  loc_00DF7394: lea ecx, arg_10
  loc_00DF7398: mov eax, [esi]
  loc_00DF739A: push edx
  loc_00DF739B: push ecx
  loc_00DF739C: push esi
  loc_00DF739D: call [eax+0000097Ch]
  loc_00DF73A3: mov edx, [esi]
  loc_00DF73A5: lea eax, arg_24
  loc_00DF73A9: mov ecx, Me
  loc_00DF73AD: push eax
  loc_00DF73AE: push ecx
  loc_00DF73AF: push esi
  loc_00DF73B0: call [edx+00000974h]
  loc_00DF73B6: mov edx, [esi]
  loc_00DF73B8: lea eax, Me
  loc_00DF73BC: lea ecx, Proc_2_129_DF6D20
  loc_00DF73C0: push eax
  loc_00DF73C1: push ecx
  loc_00DF73C2: push 00FFFFFFh
  loc_00DF73C7: push esi
  loc_00DF73C8: mov arg_14, edi
  loc_00DF73CC: call [edx+0000090Ch]
  loc_00DF73D2: mov edx, [esi]
  loc_00DF73D4: lea eax, arg_10
  loc_00DF73D8: mov ecx, Me
  loc_00DF73DC: push eax
  loc_00DF73DD: mov eax, arg_10
  loc_00DF73E1: push ecx
  loc_00DF73E2: push eax
  loc_00DF73E3: push esi
  loc_00DF73E4: call [edx+000008ECh]
  loc_00DF73EA: mov edx, [esi]
  loc_00DF73EC: lea eax, arg_18
  loc_00DF73F0: mov ecx, arg_10
  loc_00DF73F4: push eax
  loc_00DF73F5: mov arg_20, ecx
  loc_00DF73F9: lea ecx, arg_18
  loc_00DF73FD: lea eax, arg_20
  loc_00DF7401: push ecx
  loc_00DF7402: mov ebx, 3DCCCCCDh
  loc_00DF7407: push eax
  loc_00DF7408: push esi
  loc_00DF7409: mov arg_24, ebx
  loc_00DF740D: call [edx+0000097Ch]
  loc_00DF7413: mov ecx, [esi]
  loc_00DF7415: push 00000001h
  loc_00DF7417: mov edx, arg_10
  loc_00DF741B: mov eax, arg_1C
  loc_00DF741F: push edx
  loc_00DF7420: mov edx, [esi+000001D4h]
  loc_00DF7426: push eax
  loc_00DF7427: push 00000005h
  loc_00DF7429: push edx
  loc_00DF742A: push edi
  loc_00DF742B: push edi
  loc_00DF742C: push esi
  loc_00DF742D: call [ecx+00000908h]
  loc_00DF7433: fld real8 ptr [00401590h]
  loc_00DF7439: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF743D: fnstsw ax
  loc_00DF743F: test al, 0Dh
  loc_00DF7441: jnz 00DF7576h
  loc_00DF7447: mov eax, [esi]
  loc_00DF7449: lea ecx, Me
  loc_00DF744D: push ecx
  loc_00DF744E: lea edx, Me
  loc_00DF7452: lea ecx, arg_10
  loc_00DF7456: push edx
  loc_00DF7457: push ecx
  loc_00DF7458: push esi
  loc_00DF7459: call [eax+0000097Ch]
  loc_00DF745F: mov edx, [esi]
  loc_00DF7461: lea eax, arg_1C
  loc_00DF7465: lea ecx, arg_10
  loc_00DF7469: push eax
  loc_00DF746A: push ecx
  loc_00DF746B: push 00FFFFFFh
  loc_00DF7470: push esi
  loc_00DF7471: mov arg_20, edi
  loc_00DF7475: call [edx+0000090Ch]
  loc_00DF747B: mov edx, [esi]
  loc_00DF747D: lea eax, arg_18
  loc_00DF7481: push eax
  loc_00DF7482: lea ecx, arg_18
  loc_00DF7486: lea eax, arg_10
  loc_00DF748A: push ecx
  loc_00DF748B: push eax
  loc_00DF748C: push esi
  loc_00DF748D: mov arg_24, ebx
  loc_00DF7491: call [edx+0000097Ch]
  loc_00DF7497: mov ecx, [esi]
  loc_00DF7499: lea edx, arg_20
  loc_00DF749D: mov eax, arg_18
  loc_00DF74A1: push edx
  loc_00DF74A2: mov edx, arg_20
  loc_00DF74A6: push eax
  loc_00DF74A7: push edx
  loc_00DF74A8: push esi
  loc_00DF74A9: call [ecx+000008ECh]
  loc_00DF74AF: push 00000001h
  loc_00DF74B1: mov eax, [esi]
  loc_00DF74B3: mov ecx, arg_24
  loc_00DF74B7: mov edx, arg_C
  loc_00DF74BB: push ecx
  loc_00DF74BC: mov ecx, [esi+000001D0h]
  loc_00DF74C2: sub ecx, 00000001h
  loc_00DF74C5: push edx
  loc_00DF74C6: mov edx, [esi+000001D4h]
  loc_00DF74CC: jo 00DF757Bh
  loc_00DF74D2: push ecx
  loc_00DF74D3: push edx
  loc_00DF74D4: push 00000006h
  loc_00DF74D6: push edi
  loc_00DF74D7: push esi
  loc_00DF74D8: call [eax+00000908h]
  loc_00DF74DE: mov eax, [esi]
  loc_00DF74E0: push esi
  loc_00DF74E1: call [eax+00000920h]
  loc_00DF74E7: fld real8 ptr [00401818h]
  loc_00DF74ED: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF74F1: fnstsw ax
  loc_00DF74F3: test al, 0Dh
  loc_00DF74F5: jnz 00DF7576h
  loc_00DF74F7: mov ecx, [esi]
  loc_00DF74F9: lea edx, Me
  loc_00DF74FD: push edx
  loc_00DF74FE: lea eax, Me
  loc_00DF7502: lea edx, arg_10
  loc_00DF7506: push eax
  loc_00DF7507: push edx
  loc_00DF7508: push esi
  loc_00DF7509: call [ecx+0000097Ch]
  loc_00DF750F: mov edx, [esi+000001D0h]
  loc_00DF7515: mov eax, [esi]
  loc_00DF7517: mov ecx, Me
  loc_00DF751B: push ecx
  loc_00DF751C: mov ecx, [esi+000001D4h]
  loc_00DF7522: push edx
  loc_00DF7523: push ecx
  loc_00DF7524: push edi
  loc_00DF7525: push edi
  loc_00DF7526: push esi
  loc_00DF7527: call [eax+000008F0h]
  loc_00DF752D: fld real8 ptr [00401828h]
  loc_00DF7533: fstp real4 ptr Proc_2_129_DF6D20
  loc_00DF7537: fnstsw ax
  loc_00DF7539: test al, 0Dh
  loc_00DF753B: jnz 00DF7576h
  loc_00DF753D: mov edx, [esi]
  loc_00DF753F: lea ecx, Proc_2_129_DF6D20
  loc_00DF7543: lea eax, Me
  loc_00DF7547: push eax
  loc_00DF7548: lea eax, arg_10
  loc_00DF754C: push ecx
  loc_00DF754D: push eax
  loc_00DF754E: push esi
  loc_00DF754F: call [edx+0000097Ch]
  loc_00DF7555: mov edx, [esi]
  loc_00DF7557: lea eax, arg_10
  loc_00DF755B: mov ecx, Me
  loc_00DF755F: push eax
  loc_00DF7560: push esi
  loc_00DF7561: mov arg_18, ecx
  loc_00DF7565: call [edx+00000940h]
  loc_00DF756B: pop edi
  loc_00DF756C: pop esi
  loc_00DF756D: xor eax, eax
  loc_00DF756F: pop ebx
  loc_00DF7570: add esp, 00000030h
  loc_00DF7573: retn 0008h
End Sub

Private Sub Proc_2_130_DF7590(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28) 'DF7590
  loc_00DF7590: sub esp, 00000030h
  loc_00DF7593: push ebx
  loc_00DF7594: xor eax, eax
  loc_00DF7596: push ebp
  loc_00DF7597: push esi
  loc_00DF7598: mov esi, arg_38
  loc_00DF759C: mov arg_24, eax
  loc_00DF75A0: mov arg_28, eax
  loc_00DF75A4: lea edx, Proc_2_130_DF7590
  loc_00DF75A8: mov ecx, [esi]
  loc_00DF75AA: push edi
  loc_00DF75AB: mov arg_30, eax
  loc_00DF75AF: xor edi, edi
  loc_00DF75B1: push edx
  loc_00DF75B2: push esi
  loc_00DF75B3: mov arg_3C, eax
  loc_00DF75B7: mov arg_10, edi
  loc_00DF75BB: mov arg_14, edi
  loc_00DF75BF: mov arg_18, edi
  loc_00DF75C3: mov arg_1C, edi
  loc_00DF75C7: mov arg_20, edi
  loc_00DF75CB: mov arg_24, edi
  loc_00DF75CF: call [ecx+00000108h]
  loc_00DF75D5: mov ebp, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF75DB: cmp eax, edi
  loc_00DF75DD: fnclex
  loc_00DF75DF: jge 00DF75EFh
  loc_00DF75E1: push 00000108h
  loc_00DF75E6: push 006DCB00h
  loc_00DF75EB: push esi
  loc_00DF75EC: push eax
  loc_00DF75ED: call ebp
  loc_00DF75EF: fld real4 ptr Me
  loc_00DF75F3: mov ebx, [00401208h] ; __vbaFpI4
  loc_00DF75F9: call ebx
  loc_00DF75FB: lea ecx, Me
  loc_00DF75FF: mov [esi+000001D0h], eax
  loc_00DF7605: mov eax, [esi]
  loc_00DF7607: push ecx
  loc_00DF7608: push esi
  loc_00DF7609: call [eax+00000100h]
  loc_00DF760F: cmp eax, edi
  loc_00DF7611: fnclex
  loc_00DF7613: jge 00DF7623h
  loc_00DF7615: push 00000100h
  loc_00DF761A: push 006DCB00h
  loc_00DF761F: push esi
  loc_00DF7620: push eax
  loc_00DF7621: call ebp
  loc_00DF7623: fld real4 ptr Me
  loc_00DF7627: call ebx
  loc_00DF7629: cmp [esi+0000007Ch], di
  loc_00DF762D: mov [esi+000001D4h], eax
  loc_00DF7633: jnz 00DF76B1h
  loc_00DF7635: mov edx, [esi]
  loc_00DF7637: lea eax, arg_C
  loc_00DF763B: push eax
  loc_00DF763C: mov eax, [esi+00000088h]
  loc_00DF7642: lea ecx, arg_C
  loc_00DF7646: mov arg_C, edi
  loc_00DF764A: push ecx
  loc_00DF764B: push eax
  loc_00DF764C: push esi
  loc_00DF764D: call [edx+0000090Ch]
  loc_00DF7653: mov ecx, [esi]
  loc_00DF7655: lea edx, [esi+000000ACh]
  loc_00DF765B: mov eax, arg_C
  loc_00DF765F: push edx
  loc_00DF7660: push eax
  loc_00DF7661: push esi
  loc_00DF7662: call [ecx+00000974h]
  loc_00DF7668: mov ecx, [esi]
  loc_00DF766A: push esi
  loc_00DF766B: call [ecx+00000920h]
  loc_00DF7671: mov edx, [esi]
  loc_00DF7673: lea eax, arg_C
  loc_00DF7677: push eax
  loc_00DF7678: mov eax, [esi+00000088h]
  loc_00DF767E: lea ecx, arg_C
  loc_00DF7682: mov arg_C, edi
  loc_00DF7686: push ecx
  loc_00DF7687: push eax
  loc_00DF7688: push esi
  loc_00DF7689: call [edx+0000090Ch]
  loc_00DF768F: mov edx, [esi]
  loc_00DF7691: lea eax, arg_10
  loc_00DF7695: mov ecx, arg_C
  loc_00DF7699: push eax
  loc_00DF769A: push esi
  loc_00DF769B: mov arg_18, ecx
  loc_00DF769F: call [edx+00000940h]
  loc_00DF76A5: pop edi
  loc_00DF76A6: pop esi
  loc_00DF76A7: pop ebp
  loc_00DF76A8: xor eax, eax
  loc_00DF76AA: pop ebx
  loc_00DF76AB: add esp, 00000030h
  loc_00DF76AE: retn 0008h
End Sub

Private Sub Proc_2_131_DF7A60(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28, arg_2C, arg_30, arg_34) 'DF7A60
  loc_00DF7A60: sub esp, 00000044h
  loc_00DF7A63: push ebx
  loc_00DF7A64: xor eax, eax
  loc_00DF7A66: push ebp
  loc_00DF7A67: push esi
  loc_00DF7A68: mov esi, Proc_2_131_DF7A60C
  loc_00DF7A6C: mov arg_38, eax
  loc_00DF7A70: mov arg_3C, eax
  loc_00DF7A74: lea edx, Proc_2_131_DF7A60
  loc_00DF7A78: mov ecx, [esi]
  loc_00DF7A7A: push edi
  loc_00DF7A7B: mov Proc_2_131_DF7A604, eax
  loc_00DF7A7F: xor ebx, ebx
  loc_00DF7A81: push edx
  loc_00DF7A82: push esi
  loc_00DF7A83: mov arg_50, eax
  loc_00DF7A87: mov arg_38, ebx
  loc_00DF7A8B: mov arg_10, ebx
  loc_00DF7A8F: mov arg_14, ebx
  loc_00DF7A93: mov arg_18, ebx
  loc_00DF7A97: mov arg_1C, ebx
  loc_00DF7A9B: mov arg_20, ebx
  loc_00DF7A9F: mov arg_24, ebx
  loc_00DF7AA3: mov arg_28, ebx
  loc_00DF7AA7: mov arg_30, ebx
  loc_00DF7AAB: mov arg_2C, ebx
  loc_00DF7AAF: mov arg_34, ebx
  loc_00DF7AB3: call [ecx+00000108h]
  loc_00DF7AB9: cmp eax, ebx
  loc_00DF7ABB: fnclex
  loc_00DF7ABD: jge 00DF7AD1h
  loc_00DF7ABF: push 00000108h
  loc_00DF7AC4: push 006DCB00h
  loc_00DF7AC9: push esi
  loc_00DF7ACA: push eax
  loc_00DF7ACB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF7AD1: fld real4 ptr Me
  loc_00DF7AD5: mov edi, [00401208h] ; __vbaFpI4
  loc_00DF7ADB: call edi
  loc_00DF7ADD: lea ecx, Me
  loc_00DF7AE1: mov [esi+000001D0h], eax
  loc_00DF7AE7: mov eax, [esi]
  loc_00DF7AE9: push ecx
  loc_00DF7AEA: push esi
  loc_00DF7AEB: call [eax+00000100h]
  loc_00DF7AF1: cmp eax, ebx
  loc_00DF7AF3: fnclex
  loc_00DF7AF5: jge 00DF7B09h
  loc_00DF7AF7: push 00000100h
  loc_00DF7AFC: push 006DCB00h
  loc_00DF7B01: push esi
  loc_00DF7B02: push eax
  loc_00DF7B03: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF7B09: fld real4 ptr Me
  loc_00DF7B0D: call edi
  loc_00DF7B0F: mov edx, [esi]
  loc_00DF7B11: mov [esi+000001D4h], eax
  loc_00DF7B17: lea eax, arg_C
  loc_00DF7B1B: lea ecx, Me
  loc_00DF7B1F: push eax
  loc_00DF7B20: mov eax, [esi+00000088h]
  loc_00DF7B26: push ecx
  loc_00DF7B27: push eax
  loc_00DF7B28: push esi
  loc_00DF7B29: mov arg_18, ebx
  loc_00DF7B2D: call [edx+0000090Ch]
  loc_00DF7B33: mov edx, [esi]
  loc_00DF7B35: lea eax, arg_18
  loc_00DF7B39: mov ecx, arg_C
  loc_00DF7B3D: push eax
  loc_00DF7B3E: mov arg_14, ecx
  loc_00DF7B42: lea ecx, arg_18
  loc_00DF7B46: lea eax, arg_14
  loc_00DF7B4A: push ecx
  loc_00DF7B4B: push eax
  loc_00DF7B4C: push esi
  loc_00DF7B4D: mov arg_24, 3D4CCCCDh
  loc_00DF7B55: call [edx+0000097Ch]
  loc_00DF7B5B: mov ecx, [esi]
  loc_00DF7B5D: mov ebp, arg_18
  loc_00DF7B61: lea edx, arg_C
  loc_00DF7B65: lea eax, Me
  loc_00DF7B69: push edx
  loc_00DF7B6A: mov edx, [esi+00000088h]
  loc_00DF7B70: push eax
  loc_00DF7B71: push edx
  loc_00DF7B72: push esi
  loc_00DF7B73: mov arg_18, ebx
  loc_00DF7B77: call [ecx+0000090Ch]
  loc_00DF7B7D: cmp [esi+0000007Ch], bx
  loc_00DF7B81: mov eax, arg_C
  loc_00DF7B85: mov arg_30, eax
  loc_00DF7B89: jnz 00DF7CDEh
  loc_00DF7B8F: mov ecx, [esi]
  loc_00DF7B91: push esi
  loc_00DF7B92: call [ecx+00000914h]
  loc_00DF7B98: mov edx, [esi+000001D0h]
  loc_00DF7B9E: mov eax, [esi+000001D4h]
  loc_00DF7BA4: push edx
  loc_00DF7BA5: push eax
  loc_00DF7BA6: push ebx
  loc_00DF7BA7: lea ecx, Proc_2_131_DF7A608
  loc_00DF7BAB: push ebx
  loc_00DF7BAC: push ecx
  loc_00DF7BAD: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF7BB2: call [00401074h] ; __vbaSetSystemError
  loc_00DF7BB8: mov edx, [esi]
  loc_00DF7BBA: lea eax, arg_C
  loc_00DF7BBE: push eax
  loc_00DF7BBF: lea ecx, arg_C
  loc_00DF7BC3: lea eax, arg_34
  loc_00DF7BC7: push ecx
  loc_00DF7BC8: push eax
  loc_00DF7BC9: push esi
  loc_00DF7BCA: mov arg_18, 3CF5C28Fh
  loc_00DF7BD2: call [edx+0000097Ch]
  loc_00DF7BD8: mov ecx, [esi]
  loc_00DF7BDA: lea edx, arg_3C
  loc_00DF7BDE: mov eax, arg_C
  loc_00DF7BE2: push edx
  loc_00DF7BE3: push eax
  loc_00DF7BE4: push esi
  loc_00DF7BE5: call [ecx+00000974h]
  loc_00DF7BEB: mov ecx, [esi]
  loc_00DF7BED: push esi
  loc_00DF7BEE: call [ecx+00000920h]
  loc_00DF7BF4: fld real8 ptr [00401800h]
  loc_00DF7BFA: fstp real4 ptr Me
  loc_00DF7BFE: fnstsw ax
  loc_00DF7C00: test al, 0Dh
  loc_00DF7C02: jnz 00DF8CC2h
  loc_00DF7C08: mov edx, [esi]
  loc_00DF7C0A: lea ecx, Me
  loc_00DF7C0E: lea eax, arg_C
  loc_00DF7C12: push eax
  loc_00DF7C13: lea eax, arg_34
  loc_00DF7C17: push ecx
  loc_00DF7C18: push eax
  loc_00DF7C19: push esi
  loc_00DF7C1A: call [edx+0000097Ch]
  loc_00DF7C20: mov eax, [esi+000001D0h]
  loc_00DF7C26: mov ecx, [esi]
  loc_00DF7C28: mov edx, arg_C
  loc_00DF7C2C: push edx
  loc_00DF7C2D: mov edx, [esi+000001D4h]
  loc_00DF7C33: push eax
  loc_00DF7C34: push edx
  loc_00DF7C35: push ebx
  loc_00DF7C36: push ebx
  loc_00DF7C37: push esi
  loc_00DF7C38: call [ecx+000008F0h]
  loc_00DF7C3E: mov eax, [esi]
  loc_00DF7C40: lea ecx, arg_C
  loc_00DF7C44: push ecx
  loc_00DF7C45: lea edx, arg_C
  loc_00DF7C49: lea ecx, arg_34
  loc_00DF7C4D: push edx
  loc_00DF7C4E: push ecx
  loc_00DF7C4F: push esi
  loc_00DF7C50: mov arg_18, 3E800000h
  loc_00DF7C58: call [eax+0000097Ch]
  loc_00DF7C5E: mov ecx, [esi+000001D0h]
  loc_00DF7C64: mov edx, [esi]
  loc_00DF7C66: mov eax, arg_C
  loc_00DF7C6A: sub ecx, 00000002h
  loc_00DF7C6D: push eax
  loc_00DF7C6E: mov eax, [esi+000001D4h]
  loc_00DF7C74: jo 00DF8CC7h
  loc_00DF7C7A: sub eax, 00000002h
  loc_00DF7C7D: push ecx
  loc_00DF7C7E: jo 00DF8CC7h
  loc_00DF7C84: push eax
  loc_00DF7C85: push 00000001h
  loc_00DF7C87: push 00000001h
  loc_00DF7C89: push esi
  loc_00DF7C8A: call [edx+000008F0h]
  loc_00DF7C90: fld real8 ptr [00401830h]
  loc_00DF7C96: fstp real4 ptr Me
  loc_00DF7C9A: fnstsw ax
  loc_00DF7C9C: test al, 0Dh
  loc_00DF7C9E: jnz 00DF8CC2h
  loc_00DF7CA4: mov ecx, [esi]
  loc_00DF7CA6: lea edx, arg_C
  loc_00DF7CAA: push edx
  loc_00DF7CAB: lea eax, arg_C
  loc_00DF7CAF: lea edx, arg_34
  loc_00DF7CB3: push eax
  loc_00DF7CB4: push edx
  loc_00DF7CB5: push esi
  loc_00DF7CB6: call [ecx+0000097Ch]
  loc_00DF7CBC: mov ecx, [esi]
  loc_00DF7CBE: lea edx, arg_10
  loc_00DF7CC2: mov eax, arg_C
  loc_00DF7CC6: push edx
  loc_00DF7CC7: push esi
  loc_00DF7CC8: mov arg_18, eax
  loc_00DF7CCC: call [ecx+00000940h]
  loc_00DF7CD2: pop edi
  loc_00DF7CD3: pop esi
  loc_00DF7CD4: pop ebp
  loc_00DF7CD5: xor eax, eax
  loc_00DF7CD7: pop ebx
  loc_00DF7CD8: add esp, 00000044h
  loc_00DF7CDB: retn 0008h
End Sub

Private Sub Proc_2_132_DF8CD0(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28, arg_2C, arg_30, arg_34) 'DF8CD0
  loc_00DF8CD0: sub esp, 0000002Ch
  loc_00DF8CD3: push ebx
  loc_00DF8CD4: push ebp
  loc_00DF8CD5: push esi
  loc_00DF8CD6: mov esi, arg_34
  loc_00DF8CDA: lea ecx, Proc_2_132_DF8CD0
  loc_00DF8CDE: push edi
  loc_00DF8CDF: mov eax, [esi]
  loc_00DF8CE1: xor ebx, ebx
  loc_00DF8CE3: push ecx
  loc_00DF8CE4: push esi
  loc_00DF8CE5: mov arg_20, ebx
  loc_00DF8CE9: mov arg_10, ebx
  loc_00DF8CED: mov arg_14, ebx
  loc_00DF8CF1: mov arg_18, ebx
  loc_00DF8CF5: mov arg_1C, ebx
  loc_00DF8CF9: mov arg_30, ebx
  loc_00DF8CFD: mov arg_24, ebx
  loc_00DF8D01: mov arg_28, ebx
  loc_00DF8D05: mov arg_2C, ebx
  loc_00DF8D09: call [eax+00000108h]
  loc_00DF8D0F: mov ebp, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF8D15: cmp eax, ebx
  loc_00DF8D17: fnclex
  loc_00DF8D19: jge 00DF8D29h
  loc_00DF8D1B: push 00000108h
  loc_00DF8D20: push 006DCB00h
  loc_00DF8D25: push esi
  loc_00DF8D26: push eax
  loc_00DF8D27: call ebp
  loc_00DF8D29: fld real4 ptr Me
  loc_00DF8D2D: mov edi, [00401208h] ; __vbaFpI4
  loc_00DF8D33: call edi
  loc_00DF8D35: mov edx, [esi]
  loc_00DF8D37: mov [esi+000001D0h], eax
  loc_00DF8D3D: lea eax, Me
  loc_00DF8D41: push eax
  loc_00DF8D42: push esi
  loc_00DF8D43: call [edx+00000100h]
  loc_00DF8D49: cmp eax, ebx
  loc_00DF8D4B: fnclex
  loc_00DF8D4D: jge 00DF8D5Dh
  loc_00DF8D4F: push 00000100h
  loc_00DF8D54: push 006DCB00h
  loc_00DF8D59: push esi
  loc_00DF8D5A: push eax
  loc_00DF8D5B: call ebp
  loc_00DF8D5D: fld real4 ptr Me
  loc_00DF8D61: call edi
  loc_00DF8D63: mov ecx, [esi]
  loc_00DF8D65: lea edx, arg_C
  loc_00DF8D69: mov [esi+000001D4h], eax
  loc_00DF8D6F: push edx
  loc_00DF8D70: mov edx, [esi+00000088h]
  loc_00DF8D76: lea eax, arg_C
  loc_00DF8D7A: push eax
  loc_00DF8D7B: push edx
  loc_00DF8D7C: push esi
  loc_00DF8D7D: mov arg_18, ebx
  loc_00DF8D81: call [ecx+0000090Ch]
  loc_00DF8D87: mov eax, [esi+00000064h]
  loc_00DF8D8A: mov ecx, arg_C
  loc_00DF8D8E: cmp eax, ebx
  loc_00DF8D90: mov arg_18, ecx
  loc_00DF8D94: jz 00DF9123h
  loc_00DF8D9A: cmp [esi+0000006Ch], bx
  loc_00DF8D9E: jz 00DF9123h
  loc_00DF8DA4: mov eax, [esi]
  loc_00DF8DA6: lea ecx, arg_C
  loc_00DF8DAA: lea edx, Me
  loc_00DF8DAE: push ecx
  loc_00DF8DAF: push edx
  loc_00DF8DB0: push 00A9D9FFh
  loc_00DF8DB5: push esi
  loc_00DF8DB6: mov arg_18, ebx
  loc_00DF8DBA: call [eax+0000090Ch]
  loc_00DF8DC0: mov eax, [esi]
  loc_00DF8DC2: lea ecx, arg_14
  loc_00DF8DC6: lea edx, arg_10
  loc_00DF8DCA: push ecx
  loc_00DF8DCB: push edx
  loc_00DF8DCC: push 006FC0FFh
  loc_00DF8DD1: push esi
  loc_00DF8DD2: mov arg_20, ebx
  loc_00DF8DD6: call [eax+0000090Ch]
  loc_00DF8DDC: fild real4 ptr [esi+000001D0h]
  loc_00DF8DE2: mov eax, arg_14
  loc_00DF8DE6: mov ecx, arg_C
  loc_00DF8DEA: mov ebp, [esi]
  loc_00DF8DEC: push 00000001h
  loc_00DF8DEE: fstp real8 ptr arg_30
  loc_00DF8DF2: fld real8 ptr arg_30
  loc_00DF8DF6: cmp [00E3D000h], 00000000h
  loc_00DF8DFD: jnz 00DF8E07h
  loc_00DF8DFF: fdiv st0, real8 ptr [00401860h]
  loc_00DF8E05: jmp 00DF8E18h
  loc_00DF8E07: push [00401864h]
  loc_00DF8E0D: push [00401860h]
  loc_00DF8E13: call 00402854h ; _adj_fdiv_m64
  loc_00DF8E18: push eax
  loc_00DF8E19: push ecx
  loc_00DF8E1A: fnstsw ax
  loc_00DF8E1C: test al, 0Dh
  loc_00DF8E1E: jnz 00DF9640h
  loc_00DF8E24: call edi
  loc_00DF8E26: mov edx, [esi+000001D4h]
  loc_00DF8E2C: push eax
  loc_00DF8E2D: push edx
  loc_00DF8E2E: push ebx
  loc_00DF8E2F: push ebx
  loc_00DF8E30: push esi
  loc_00DF8E31: call arg_908
  loc_00DF8E37: mov eax, [esi]
  loc_00DF8E39: lea ecx, arg_C
  loc_00DF8E3D: lea edx, Me
  loc_00DF8E41: push ecx
  loc_00DF8E42: push edx
  loc_00DF8E43: push 003FABFFh
  loc_00DF8E48: push esi
  loc_00DF8E49: mov arg_18, ebx
  loc_00DF8E4D: call [eax+0000090Ch]
  loc_00DF8E53: mov eax, [esi]
  loc_00DF8E55: lea ecx, arg_14
  loc_00DF8E59: lea edx, arg_10
  loc_00DF8E5D: push ecx
  loc_00DF8E5E: push edx
  loc_00DF8E5F: push 0075E1FFh
  loc_00DF8E64: push esi
  loc_00DF8E65: mov arg_20, ebx
  loc_00DF8E69: call [eax+0000090Ch]
  loc_00DF8E6F: fild real4 ptr [esi+000001D0h]
  loc_00DF8E75: mov eax, arg_14
  loc_00DF8E79: mov ecx, arg_C
  loc_00DF8E7D: mov ebp, [esi]
  loc_00DF8E7F: push 00000001h
  loc_00DF8E81: fstp real8 ptr arg_30
  loc_00DF8E85: fld real8 ptr arg_30
  loc_00DF8E89: cmp [00E3D000h], 00000000h
  loc_00DF8E90: jnz 00DF8E9Ah
  loc_00DF8E92: fdiv st0, real8 ptr [00401860h]
  loc_00DF8E98: jmp 00DF8EABh
  loc_00DF8E9A: push [00401864h]
  loc_00DF8EA0: push [00401860h]
  loc_00DF8EA6: call 00402854h ; _adj_fdiv_m64
  loc_00DF8EAB: push eax
  loc_00DF8EAC: push ecx
  loc_00DF8EAD: fsubr st0, real8 ptr arg_38
  loc_00DF8EB1: fnstsw ax
  loc_00DF8EB3: test al, 0Dh
  loc_00DF8EB5: jnz 00DF9640h
  loc_00DF8EBB: call edi
  loc_00DF8EBD: fild real4 ptr [esi+000001D0h]
  loc_00DF8EC3: mov edx, [esi+000001D4h]
  loc_00DF8EC9: push eax
  loc_00DF8ECA: push edx
  loc_00DF8ECB: fstp real8 ptr Proc_2_132_DF8CD00
  loc_00DF8ECF: fld real8 ptr Proc_2_132_DF8CD00
  loc_00DF8ED3: cmp [00E3D000h], 00000000h
  loc_00DF8EDA: jnz 00DF8EE4h
  loc_00DF8EDC: fdiv st0, real8 ptr [00401860h]
  loc_00DF8EE2: jmp 00DF8EF5h
  loc_00DF8EE4: push [00401864h]
  loc_00DF8EEA: push [00401860h]
  loc_00DF8EF0: call 00402854h ; _adj_fdiv_m64
  loc_00DF8EF5: fnstsw ax
  loc_00DF8EF7: test al, 0Dh
  loc_00DF8EF9: jnz 00DF9640h
  loc_00DF8EFF: call edi
  loc_00DF8F01: push eax
  loc_00DF8F02: push ebx
  loc_00DF8F03: push esi
  loc_00DF8F04: call arg_908
  loc_00DF8F0A: fld real8 ptr [00401858h]
  loc_00DF8F10: fstp real4 ptr Me
  loc_00DF8F14: fnstsw ax
  loc_00DF8F16: test al, 0Dh
  loc_00DF8F18: jnz 00DF9640h
  loc_00DF8F1E: lea ecx, arg_C
  loc_00DF8F22: lea edx, Me
  loc_00DF8F26: push ecx
  loc_00DF8F27: lea ecx, arg_1C
  loc_00DF8F2B: mov eax, [esi]
  loc_00DF8F2D: push edx
  loc_00DF8F2E: push ecx
  loc_00DF8F2F: push esi
  loc_00DF8F30: call [eax+0000097Ch]
  loc_00DF8F36: mov ecx, [esi+000001D0h]
  loc_00DF8F3C: mov edx, [esi]
  loc_00DF8F3E: mov eax, arg_C
  loc_00DF8F42: push eax
  loc_00DF8F43: mov eax, [esi+000001D4h]
  loc_00DF8F49: push ecx
  loc_00DF8F4A: push eax
  loc_00DF8F4B: push ebx
  loc_00DF8F4C: push ebx
  loc_00DF8F4D: push esi
  loc_00DF8F4E: call [edx+000008F0h]
  loc_00DF8F54: cmp [esi+0000004Eh], bx
  loc_00DF8F58: jz 00DF910Eh
  loc_00DF8F5E: mov ecx, [esi]
  loc_00DF8F60: lea edx, arg_C
  loc_00DF8F64: lea eax, Me
  loc_00DF8F68: push edx
  loc_00DF8F69: push eax
  loc_00DF8F6A: push 0058C1FFh
  loc_00DF8F6F: push esi
  loc_00DF8F70: mov arg_18, ebx
  loc_00DF8F74: call [ecx+0000090Ch]
  loc_00DF8F7A: mov ecx, [esi]
  loc_00DF8F7C: lea edx, arg_14
  loc_00DF8F80: lea eax, arg_10
  loc_00DF8F84: push edx
  loc_00DF8F85: push eax
  loc_00DF8F86: push 0051AFFFh
  loc_00DF8F8B: push esi
  loc_00DF8F8C: mov arg_20, ebx
  loc_00DF8F90: call [ecx+0000090Ch]
  loc_00DF8F96: fild real4 ptr [esi+000001D0h]
  loc_00DF8F9C: mov ecx, arg_14
  loc_00DF8FA0: mov edx, arg_C
  loc_00DF8FA4: mov ebp, [esi]
  loc_00DF8FA6: push 00000001h
  loc_00DF8FA8: fstp real8 ptr arg_30
  loc_00DF8FAC: fld real8 ptr arg_30
  loc_00DF8FB0: cmp [00E3D000h], 00000000h
  loc_00DF8FB7: jnz 00DF8FC1h
  loc_00DF8FB9: fdiv st0, real8 ptr [00401860h]
  loc_00DF8FBF: jmp 00DF8FD2h
  loc_00DF8FC1: push [00401864h]
  loc_00DF8FC7: push [00401860h]
  loc_00DF8FCD: call 00402854h ; _adj_fdiv_m64
  loc_00DF8FD2: push ecx
  loc_00DF8FD3: push edx
  loc_00DF8FD4: fnstsw ax
  loc_00DF8FD6: test al, 0Dh
  loc_00DF8FD8: jnz 00DF9640h
  loc_00DF8FDE: call edi
  loc_00DF8FE0: push eax
  loc_00DF8FE1: mov eax, [esi+000001D4h]
  loc_00DF8FE7: push eax
  loc_00DF8FE8: push ebx
  loc_00DF8FE9: push ebx
  loc_00DF8FEA: push esi
  loc_00DF8FEB: call arg_908
  loc_00DF8FF1: mov ecx, [esi]
  loc_00DF8FF3: lea edx, arg_C
  loc_00DF8FF7: lea eax, Me
  loc_00DF8FFB: push edx
  loc_00DF8FFC: push eax
  loc_00DF8FFD: push 00468FFFh
  loc_00DF9002: push esi
  loc_00DF9003: mov arg_18, ebx
  loc_00DF9007: call [ecx+0000090Ch]
  loc_00DF900D: mov ecx, [esi]
  loc_00DF900F: lea edx, arg_14
  loc_00DF9013: lea eax, arg_10
  loc_00DF9017: push edx
  loc_00DF9018: push eax
  loc_00DF9019: push 005FD3FFh
  loc_00DF901E: push esi
  loc_00DF901F: mov arg_20, ebx
  loc_00DF9023: call [ecx+0000090Ch]
  loc_00DF9029: fild real4 ptr [esi+000001D0h]
  loc_00DF902F: mov ecx, arg_14
  loc_00DF9033: mov edx, arg_C
  loc_00DF9037: mov ebp, [esi]
  loc_00DF9039: push 00000001h
  loc_00DF903B: fstp real8 ptr arg_30
  loc_00DF903F: fld real8 ptr arg_30
  loc_00DF9043: cmp [00E3D000h], 00000000h
  loc_00DF904A: jnz 00DF9054h
  loc_00DF904C: fdiv st0, real8 ptr [00401860h]
  loc_00DF9052: jmp 00DF9065h
  loc_00DF9054: push [00401864h]
  loc_00DF905A: push [00401860h]
  loc_00DF9060: call 00402854h ; _adj_fdiv_m64
  loc_00DF9065: push ecx
  loc_00DF9066: push edx
  loc_00DF9067: fsubr st0, real8 ptr arg_38
  loc_00DF906B: fnstsw ax
  loc_00DF906D: test al, 0Dh
  loc_00DF906F: jnz 00DF9640h
  loc_00DF9075: call edi
  loc_00DF9077: fild real4 ptr [esi+000001D0h]
  loc_00DF907D: push eax
  loc_00DF907E: mov eax, [esi+000001D4h]
  loc_00DF9084: push eax
  loc_00DF9085: fstp real8 ptr Proc_2_132_DF8CD00
  loc_00DF9089: fld real8 ptr Proc_2_132_DF8CD00
  loc_00DF908D: cmp [00E3D000h], 00000000h
  loc_00DF9094: jnz 00DF909Eh
  loc_00DF9096: fdiv st0, real8 ptr [00401860h]
  loc_00DF909C: jmp 00DF90AFh
  loc_00DF909E: push [00401864h]
  loc_00DF90A4: push [00401860h]
  loc_00DF90AA: call 00402854h ; _adj_fdiv_m64
  loc_00DF90AF: fnstsw ax
  loc_00DF90B1: test al, 0Dh
  loc_00DF90B3: jnz 00DF9640h
  loc_00DF90B9: call edi
  loc_00DF90BB: push eax
  loc_00DF90BC: push ebx
  loc_00DF90BD: push esi
  loc_00DF90BE: call arg_908
  loc_00DF90C4: fld real8 ptr [00401858h]
  loc_00DF90CA: fstp real4 ptr Me
  loc_00DF90CE: fnstsw ax
  loc_00DF90D0: test al, 0Dh
  loc_00DF90D2: jnz 00DF9640h
  loc_00DF90D8: mov ecx, [esi]
  loc_00DF90DA: lea edx, arg_C
  loc_00DF90DE: push edx
  loc_00DF90DF: lea eax, arg_C
  loc_00DF90E3: lea edx, arg_1C
  loc_00DF90E7: push eax
  loc_00DF90E8: push edx
  loc_00DF90E9: push esi
  loc_00DF90EA: call [ecx+0000097Ch]
  loc_00DF90F0: mov edx, [esi+000001D0h]
  loc_00DF90F6: mov eax, [esi]
  loc_00DF90F8: mov ecx, arg_C
  loc_00DF90FC: push ecx
  loc_00DF90FD: mov ecx, [esi+000001D4h]
  loc_00DF9103: push edx
  loc_00DF9104: push ecx
  loc_00DF9105: push ebx
  loc_00DF9106: push ebx
  loc_00DF9107: push esi
  loc_00DF9108: call [eax+000008F0h]
  loc_00DF910E: mov edx, [esi]
  loc_00DF9110: push esi
  loc_00DF9111: call [edx+00000920h]
  loc_00DF9117: pop edi
  loc_00DF9118: pop esi
  loc_00DF9119: pop ebp
  loc_00DF911A: xor eax, eax
  loc_00DF911C: pop ebx
  loc_00DF911D: add esp, 0000002Ch
  loc_00DF9120: retn 0008h
End Sub

Private Sub Proc_2_133_DF9650(arg_C, arg_10, arg_14, arg_18, arg_1C) 'DF9650
  loc_00DF9650: sub esp, 00000020h
  loc_00DF9653: push ebx
  loc_00DF9654: push ebp
  loc_00DF9655: push esi
  loc_00DF9656: mov esi, arg_28
  loc_00DF965A: push edi
  loc_00DF965B: xor edi, edi
  loc_00DF965D: mov eax, [esi+00000010h]
  loc_00DF9660: lea edx, Me
  loc_00DF9664: mov arg_18, edi
  loc_00DF9668: mov Me, edi
  loc_00DF966C: mov arg_C, edi
  loc_00DF9670: mov arg_10, edi
  loc_00DF9674: mov arg_14, edi
  loc_00DF9678: mov arg_1C, edi
  loc_00DF967C: mov ecx, [eax]
  loc_00DF967E: push edx
  loc_00DF967F: push eax
  loc_00DF9680: call [ecx+00000108h]
  loc_00DF9686: mov ebp, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF968C: cmp eax, edi
  loc_00DF968E: fnclex
  loc_00DF9690: jge 00DF96A3h
  loc_00DF9692: mov ecx, [esi+00000010h]
  loc_00DF9695: push 00000108h
  loc_00DF969A: push 006DCB00h
  loc_00DF969F: push ecx
  loc_00DF96A0: push eax
  loc_00DF96A1: call ebp
  loc_00DF96A3: fld real4 ptr Me
  loc_00DF96A7: mov ebx, [00401208h] ; __vbaFpI4
  loc_00DF96AD: call ebx
  loc_00DF96AF: mov [esi+000001D0h], eax
  loc_00DF96B5: mov eax, [esi+00000010h]
  loc_00DF96B8: lea ecx, Me
  loc_00DF96BC: mov edx, [eax]
  loc_00DF96BE: push ecx
  loc_00DF96BF: push eax
  loc_00DF96C0: call [edx+00000100h]
  loc_00DF96C6: cmp eax, edi
  loc_00DF96C8: fnclex
  loc_00DF96CA: jge 00DF96DDh
  loc_00DF96CC: mov edx, [esi+00000010h]
  loc_00DF96CF: push 00000100h
  loc_00DF96D4: push 006DCB00h
  loc_00DF96D9: push edx
  loc_00DF96DA: push eax
  loc_00DF96DB: call ebp
  loc_00DF96DD: fld real4 ptr Me
  loc_00DF96E1: call ebx
  loc_00DF96E3: lea ecx, arg_C
  loc_00DF96E7: lea edx, Me
  loc_00DF96EB: push ecx
  loc_00DF96EC: mov ecx, [esi+00000088h]
  loc_00DF96F2: mov [esi+000001D4h], eax
  loc_00DF96F8: mov eax, [esi]
  loc_00DF96FA: push edx
  loc_00DF96FB: push ecx
  loc_00DF96FC: push esi
  loc_00DF96FD: mov arg_18, edi
  loc_00DF9701: call [eax+0000090Ch]
  loc_00DF9707: mov eax, [esi+000001D0h]
  loc_00DF970D: mov ecx, [esi+000001D4h]
  loc_00DF9713: mov edx, arg_C
  loc_00DF9717: push eax
  loc_00DF9718: mov arg_1C, edx
  loc_00DF971C: push ecx
  loc_00DF971D: push edi
  loc_00DF971E: lea edx, [esi+000000ACh]
  loc_00DF9724: push edi
  loc_00DF9725: push edx
  loc_00DF9726: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF972B: call [00401074h] ; __vbaSetSystemError
  loc_00DF9731: cmp [esi+00000064h], edi
  loc_00DF9734: jz 00DF9819h
  loc_00DF973A: cmp [esi+0000006Ch], di
  loc_00DF973E: jz 00DF9819h
  loc_00DF9744: cmp [esi+0000004Eh], di
  loc_00DF9748: mov eax, [esi]
  loc_00DF974A: lea ecx, arg_C
  loc_00DF974E: lea edx, Me
  loc_00DF9752: push ecx
  loc_00DF9753: mov arg_C, edi
  loc_00DF9757: push edx
  loc_00DF9758: jz 00DF977Dh
  loc_00DF975A: push 004E91FEh ; "J1" & Chr(9) & "I.ø "
  loc_00DF975F: push esi
  loc_00DF9760: call [eax+0000090Ch]
  loc_00DF9766: mov eax, [esi]
  loc_00DF9768: lea ecx, arg_14
  loc_00DF976C: lea edx, arg_10
  loc_00DF9770: push ecx
  loc_00DF9771: push edx
  loc_00DF9772: mov arg_18, edi
  loc_00DF9776: push 008ED3FFh
  loc_00DF977B: jmp 00DF979Eh
  loc_00DF977D: push 008CD5FFh
  loc_00DF9782: push esi
  loc_00DF9783: call [eax+0000090Ch]
  loc_00DF9789: mov eax, [esi]
  loc_00DF978B: lea ecx, arg_14
  loc_00DF978F: lea edx, arg_10
  loc_00DF9793: push ecx
  loc_00DF9794: push edx
  loc_00DF9795: mov arg_18, edi
  loc_00DF9799: push 0055ADFFh
  loc_00DF979E: push esi
  loc_00DF979F: call [eax+0000090Ch]
  loc_00DF97A5: mov eax, [esi]
  loc_00DF97A7: push 00000001h
  loc_00DF97A9: mov ecx, arg_18
  loc_00DF97AD: mov edx, arg_10
  loc_00DF97B1: push ecx
  loc_00DF97B2: mov ecx, [esi+000001D0h]
  loc_00DF97B8: push edx
  loc_00DF97B9: mov edx, [esi+000001D4h]
  loc_00DF97BF: push ecx
  loc_00DF97C0: push edx
  loc_00DF97C1: push edi
  loc_00DF97C2: push edi
  loc_00DF97C3: push esi
  loc_00DF97C4: call [eax+00000908h]
  loc_00DF97CA: mov eax, [esi]
  loc_00DF97CC: push esi
  loc_00DF97CD: call [eax+00000920h]
  loc_00DF97D3: mov ecx, [esi]
  loc_00DF97D5: lea edx, arg_C
  loc_00DF97D9: lea eax, Me
  loc_00DF97DD: push edx
  loc_00DF97DE: push eax
  loc_00DF97DF: push 00800000h
  loc_00DF97E4: push esi
  loc_00DF97E5: mov arg_18, edi
  loc_00DF97E9: call [ecx+0000090Ch]
  loc_00DF97EF: mov eax, [esi+000001D0h]
  loc_00DF97F5: mov ecx, [esi]
  loc_00DF97F7: mov edx, arg_C
  loc_00DF97FB: push edx
  loc_00DF97FC: mov edx, [esi+000001D4h]
  loc_00DF9802: push eax
  loc_00DF9803: push edx
  loc_00DF9804: push edi
  loc_00DF9805: push edi
  loc_00DF9806: push esi
  loc_00DF9807: call [ecx+000008F0h]
  loc_00DF980D: pop edi
  loc_00DF980E: pop esi
  loc_00DF980F: pop ebp
  loc_00DF9810: xor eax, eax
  loc_00DF9812: pop ebx
  loc_00DF9813: add esp, 00000020h
  loc_00DF9816: retn 0008h
End Sub

Private Sub Proc_2_134_DF9AD0(arg_C) 'DF9AD0
  loc_00DF9AD0: push ebp
  loc_00DF9AD1: mov ebp, esp
  loc_00DF9AD3: sub esp, 00000008h
  loc_00DF9AD6: push 00402836h ; __vbaExceptHandler
  loc_00DF9ADB: mov eax, fs:[00000000h]
  loc_00DF9AE1: push eax
  loc_00DF9AE2: mov fs:[00000000h], esp
  loc_00DF9AE9: sub esp, 0000001Ch
  loc_00DF9AEC: push ebx
  loc_00DF9AED: push esi
  loc_00DF9AEE: push edi
  loc_00DF9AEF: mov var_8, esp
  loc_00DF9AF2: mov var_4, 00401868h
  loc_00DF9AF9: xor ebx, ebx
  loc_00DF9AFB: push 0000000Fh
  loc_00DF9AFD: mov var_18, ebx
  loc_00DF9B00: mov var_1C, ebx
  loc_00DF9B03: call 006DDDB4h ; GetSysColor(%x1v)
  loc_00DF9B08: mov edi, eax
  loc_00DF9B0A: call [00401074h] ; __vbaSetSystemError
  loc_00DF9B10: mov esi, Me
  loc_00DF9B13: push edi
  loc_00DF9B14: mov eax, [esi+00000010h]
  loc_00DF9B17: push eax
  loc_00DF9B18: mov ecx, [eax]
  loc_00DF9B1A: call [ecx+00000064h]
  loc_00DF9B1D: test eax, eax
  loc_00DF9B1F: fnclex
  loc_00DF9B21: jge 00DF9B35h
  loc_00DF9B23: mov edx, [esi+00000010h]
  loc_00DF9B26: push 00000064h
  loc_00DF9B28: push 006DCB00h
  loc_00DF9B2D: push edx
  loc_00DF9B2E: push eax
  loc_00DF9B2F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF9B35: cmp [esi+0000007Ch], 0000h
  loc_00DF9B3A: jnz 00DF9B7Bh
  loc_00DF9B3C: mov edx, 006DEA84h ; "Button"
  loc_00DF9B41: lea ecx, var_18
  loc_00DF9B44: call [004011B0h] ; __vbaStrCopy
  loc_00DF9B4A: mov eax, [esi]
  loc_00DF9B4C: lea ecx, var_1C
  loc_00DF9B4F: push ecx
  loc_00DF9B50: push 00000004h
  loc_00DF9B52: lea edx, var_18
  loc_00DF9B55: push 00000001h
  loc_00DF9B57: push edx
  loc_00DF9B58: push esi
  loc_00DF9B59: call [eax+00000970h]
  loc_00DF9B5F: lea ecx, var_18
  loc_00DF9B62: call [00401258h] ; __vbaFreeStr
  loc_00DF9B68: mov eax, [esi]
  loc_00DF9B6A: push esi
  loc_00DF9B6B: call [eax+00000920h]
  loc_00DF9B71: push 00DF9C03h
  loc_00DF9B76: jmp 00DF9C02h
  loc_00DF9B7B: mov eax, arg_C
  loc_00DF9B7E: sub eax, 00000000h
  loc_00DF9B81: jz 00DF9B97h
  loc_00DF9B83: dec eax
  loc_00DF9B84: jz 00DF9B90h
  loc_00DF9B86: dec eax
  loc_00DF9B87: jnz 00DF9B9Ch
  loc_00DF9B89: mov ebx, 00000003h
  loc_00DF9B8E: jmp 00DF9B9Ch
  loc_00DF9B90: mov ebx, 00000002h
  loc_00DF9B95: jmp 00DF9B9Ch
  loc_00DF9B97: mov ebx, 00000001h
  loc_00DF9B9C: mov eax, [esi+00000048h]
  loc_00DF9B9F: test eax, eax
  loc_00DF9BA1: jnz 00DF9BBDh
  loc_00DF9BA3: cmp [esi+00000050h], 0000h
  loc_00DF9BA8: jnz 00DF9BB1h
  loc_00DF9BAA: cmp [esi+00000058h], 0000h
  loc_00DF9BAF: jz 00DF9BBDh
  loc_00DF9BB1: cmp [esi+00000070h], 0000h
  loc_00DF9BB6: jz 00DF9BBDh
  loc_00DF9BB8: mov ebx, 00000005h
  loc_00DF9BBD: mov edx, 006DEA84h ; "Button"
  loc_00DF9BC2: lea ecx, var_18
  loc_00DF9BC5: call [004011B0h] ; __vbaStrCopy
  loc_00DF9BCB: mov ecx, [esi]
  loc_00DF9BCD: lea edx, var_1C
  loc_00DF9BD0: push edx
  loc_00DF9BD1: push ebx
  loc_00DF9BD2: lea eax, var_18
  loc_00DF9BD5: push 00000001h
  loc_00DF9BD7: push eax
  loc_00DF9BD8: push esi
  loc_00DF9BD9: call [ecx+00000970h]
  loc_00DF9BDF: lea ecx, var_18
  loc_00DF9BE2: call [00401258h] ; __vbaFreeStr
  loc_00DF9BE8: mov ecx, [esi]
  loc_00DF9BEA: push esi
  loc_00DF9BEB: call [ecx+00000920h]
  loc_00DF9BF1: push 00DF9C03h
  loc_00DF9BF6: jmp 00DF9C02h
  loc_00DF9BF8: lea ecx, var_18
  loc_00DF9BFB: call [00401258h] ; __vbaFreeStr
  loc_00DF9C01: ret
  loc_00DF9C02: ret
  loc_00DF9C03: mov ecx, var_10
  loc_00DF9C06: pop edi
  loc_00DF9C07: pop esi
  loc_00DF9C08: xor eax, eax
  loc_00DF9C0A: mov fs:[00000000h], ecx
  loc_00DF9C11: pop ebx
  loc_00DF9C12: mov esp, ebp
  loc_00DF9C14: pop ebp
  loc_00DF9C15: retn 0008h
End Sub

Private Sub Proc_2_135_DF9C20(arg_C, arg_10, arg_14, arg_18, arg_1C, arg_20, arg_24, arg_28, arg_2C, arg_30, arg_34) 'DF9C20
  loc_00DF9C20: sub esp, 00000018h
  loc_00DF9C23: xor eax, eax
  loc_00DF9C25: push ebx
  loc_00DF9C26: mov Proc_2_135_DF9C20, eax
  loc_00DF9C2A: push ebp
  loc_00DF9C2B: mov arg_C, eax
  loc_00DF9C2F: push esi
  loc_00DF9C30: mov esi, arg_20
  loc_00DF9C34: mov arg_14, eax
  loc_00DF9C38: mov arg_18, eax
  loc_00DF9C3C: xor ebp, ebp
  loc_00DF9C3E: mov eax, [esi+00000010h]
  loc_00DF9C41: lea edx, Proc_2_135_DF9C20
  loc_00DF9C45: push edi
  loc_00DF9C46: mov arg_C, ebp
  loc_00DF9C4A: mov Me, ebp
  loc_00DF9C4E: mov ecx, [eax]
  loc_00DF9C50: push edx
  loc_00DF9C51: push eax
  loc_00DF9C52: call [ecx+00000058h]
  loc_00DF9C55: cmp eax, ebp
  loc_00DF9C57: fnclex
  loc_00DF9C59: jge 00DF9C6Dh
  loc_00DF9C5B: mov ecx, [esi+00000010h]
  loc_00DF9C5E: push 00000058h
  loc_00DF9C60: push 006DCB00h
  loc_00DF9C65: push ecx
  loc_00DF9C66: push eax
  loc_00DF9C67: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF9C6D: mov edx, arg_28
  loc_00DF9C71: mov eax, [edx]
  loc_00DF9C73: push eax
  loc_00DF9C74: call [00401190h] ; VarPtr
  loc_00DF9C7A: push eax
  loc_00DF9C7B: mov ecx, arg_C
  loc_00DF9C7F: push ecx
  loc_00DF9C80: call 006DE018h ; OpenThemeData()
  loc_00DF9C85: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DF9C8B: mov ebx, eax
  loc_00DF9C8D: mov arg_24, ebx
  loc_00DF9C91: call edi
  loc_00DF9C93: cmp ebx, ebp
  loc_00DF9C95: jz 00DF9DD4h
  loc_00DF9C9B: mov edx, [esi+000000B8h]
  loc_00DF9CA1: mov eax, [esi+000000B4h]
  loc_00DF9CA7: mov ecx, [esi+000000B0h]
  loc_00DF9CAD: add edx, 00000002h
  loc_00DF9CB0: jo 00DF9DE7h
  loc_00DF9CB6: add eax, 00000001h
  loc_00DF9CB9: lea ebx, [esi+000000ACh]
  loc_00DF9CBF: jo 00DF9DE7h
  loc_00DF9CC5: push edx
  loc_00DF9CC6: mov edx, [ebx]
  loc_00DF9CC8: sub ecx, 00000001h
  loc_00DF9CCB: push eax
  loc_00DF9CCC: jo 00DF9DE7h
  loc_00DF9CD2: sub edx, 00000001h
  loc_00DF9CD5: push ecx
  loc_00DF9CD6: jo 00DF9DE7h
  loc_00DF9CDC: lea eax, arg_1C
  loc_00DF9CE0: push edx
  loc_00DF9CE1: push eax
  loc_00DF9CE2: call 006DDB50h ; SetRect(%x1v, %x2v, %x3v, %x4v, %x5v)
  loc_00DF9CE7: call edi
  loc_00DF9CE9: mov ecx, [esi]
  loc_00DF9CEB: lea edx, Me
  loc_00DF9CEF: push edx
  loc_00DF9CF0: push esi
  loc_00DF9CF1: call [ecx+000000D8h]
  loc_00DF9CF7: cmp eax, ebp
  loc_00DF9CF9: fnclex
  loc_00DF9CFB: jge 00DF9D0Fh
  loc_00DF9CFD: push 000000D8h
  loc_00DF9D02: push 006DCB00h
  loc_00DF9D07: push esi
  loc_00DF9D08: push eax
  loc_00DF9D09: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF9D0F: mov ebp, arg_30
  loc_00DF9D13: mov edx, arg_2C
  loc_00DF9D17: lea eax, arg_C
  loc_00DF9D1B: lea ecx, arg_10
  loc_00DF9D1F: push eax
  loc_00DF9D20: mov eax, arg_C
  loc_00DF9D24: push ecx
  loc_00DF9D25: mov ecx, arg_2C
  loc_00DF9D29: push ebp
  loc_00DF9D2A: push edx
  loc_00DF9D2B: push eax
  loc_00DF9D2C: push ecx
  loc_00DF9D2D: call 006DE108h ; GetThemeBackgroundRegion()
  loc_00DF9D32: call edi
  loc_00DF9D34: mov edx, [esi]
  loc_00DF9D36: lea eax, Me
  loc_00DF9D3A: push eax
  loc_00DF9D3B: push esi
  loc_00DF9D3C: call [edx+00000818h]
  loc_00DF9D42: test eax, eax
  loc_00DF9D44: jge 00DF9D58h
  loc_00DF9D46: push 00000818h
  loc_00DF9D4B: push 006DD090h
  loc_00DF9D50: push esi
  loc_00DF9D51: push eax
  loc_00DF9D52: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF9D58: mov ecx, arg_C
  loc_00DF9D5C: mov edx, Me
  loc_00DF9D60: push FFFFFFFFh
  loc_00DF9D62: push ecx
  loc_00DF9D63: push edx
  loc_00DF9D64: call 006DDCD8h ; SetWindowRgn(%x1v, %x2v, %x3v)
  loc_00DF9D69: call edi
  loc_00DF9D6B: mov eax, arg_C
  loc_00DF9D6F: push eax
  loc_00DF9D70: call 006DD498h ; DeleteObject(%x1v)
  loc_00DF9D75: call edi
  loc_00DF9D77: mov ecx, [esi]
  loc_00DF9D79: lea edx, Me
  loc_00DF9D7D: push edx
  loc_00DF9D7E: push esi
  loc_00DF9D7F: call [ecx+000000D8h]
  loc_00DF9D85: test eax, eax
  loc_00DF9D87: fnclex
  loc_00DF9D89: jge 00DF9D9Dh
  loc_00DF9D8B: push 000000D8h
  loc_00DF9D90: push 006DCB00h
  loc_00DF9D95: push esi
  loc_00DF9D96: push eax
  loc_00DF9D97: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF9D9D: mov eax, arg_2C
  loc_00DF9DA1: mov ecx, Me
  loc_00DF9DA5: mov edx, arg_24
  loc_00DF9DA9: push ebx
  loc_00DF9DAA: push ebx
  loc_00DF9DAB: push ebp
  loc_00DF9DAC: push eax
  loc_00DF9DAD: push ecx
  loc_00DF9DAE: push edx
  loc_00DF9DAF: call 006DE0B4h ; DrawThemeBackground()
  loc_00DF9DB4: mov esi, eax
  loc_00DF9DB6: call edi
  loc_00DF9DB8: mov ecx, arg_34
  loc_00DF9DBC: xor eax, eax
  loc_00DF9DBE: test esi, esi
  loc_00DF9DC0: setnz al
  loc_00DF9DC3: neg eax
  loc_00DF9DC5: pop edi
  loc_00DF9DC6: pop esi
  loc_00DF9DC7: mov [ecx], ax
  loc_00DF9DCA: pop ebp
  loc_00DF9DCB: xor eax, eax
  loc_00DF9DCD: pop ebx
  loc_00DF9DCE: add esp, 00000018h
  loc_00DF9DD1: retn 0014h
End Sub

Private Sub Proc_2_136_DF9DF0(arg_C, arg_10, arg_14, arg_18) 'DF9DF0
  loc_00DF9DF0: push ecx
  loc_00DF9DF1: mov eax, Proc_2_136_DF9DF0
  loc_00DF9DF5: push ebx
  loc_00DF9DF6: push ebp
  loc_00DF9DF7: push esi
  loc_00DF9DF8: push edi
  loc_00DF9DF9: push eax
  loc_00DF9DFA: mov arg_C, 00000000h
  loc_00DF9E02: call 006DD26Ch ; CreateSolidBrush(%x1v)
  loc_00DF9E07: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DF9E0D: mov Me, eax
  loc_00DF9E11: call edi
  loc_00DF9E13: mov esi, arg_10
  loc_00DF9E17: mov ebx, Me
  loc_00DF9E1B: lea edx, Me
  loc_00DF9E1F: mov ecx, [esi]
  loc_00DF9E21: push edx
  loc_00DF9E22: push esi
  loc_00DF9E23: call [ecx+000000D8h]
  loc_00DF9E29: mov ebp, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF9E2F: test eax, eax
  loc_00DF9E31: fnclex
  loc_00DF9E33: jge 00DF9E43h
  loc_00DF9E35: push 000000D8h
  loc_00DF9E3A: push 006DCB00h
  loc_00DF9E3F: push esi
  loc_00DF9E40: push eax
  loc_00DF9E41: call ebp
  loc_00DF9E43: mov eax, Me
  loc_00DF9E47: push ebx
  loc_00DF9E48: push eax
  loc_00DF9E49: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DF9E4E: mov arg_14, eax
  loc_00DF9E52: call edi
  loc_00DF9E54: mov ecx, [esi]
  loc_00DF9E56: lea edx, Me
  loc_00DF9E5A: push edx
  loc_00DF9E5B: push esi
  loc_00DF9E5C: call [ecx+000000D8h]
  loc_00DF9E62: test eax, eax
  loc_00DF9E64: fnclex
  loc_00DF9E66: jge 00DF9E76h
  loc_00DF9E68: push 000000D8h
  loc_00DF9E6D: push 006DCB00h
  loc_00DF9E72: push esi
  loc_00DF9E73: push eax
  loc_00DF9E74: call ebp
  loc_00DF9E76: mov eax, arg_18
  loc_00DF9E7A: mov ecx, Me
  loc_00DF9E7E: push ebx
  loc_00DF9E7F: push eax
  loc_00DF9E80: push ecx
  loc_00DF9E81: call 006DE278h ; FillRect(%x1v, %x2v, %x3v)
  loc_00DF9E86: call edi
  loc_00DF9E88: mov edx, [esi]
  loc_00DF9E8A: lea eax, Me
  loc_00DF9E8E: push eax
  loc_00DF9E8F: push esi
  loc_00DF9E90: call [edx+000000D8h]
  loc_00DF9E96: test eax, eax
  loc_00DF9E98: fnclex
  loc_00DF9E9A: jge 00DF9EAAh
  loc_00DF9E9C: push 000000D8h
  loc_00DF9EA1: push 006DCB00h
  loc_00DF9EA6: push esi
  loc_00DF9EA7: push eax
  loc_00DF9EA8: call ebp
  loc_00DF9EAA: mov ecx, arg_14
  loc_00DF9EAE: mov edx, Me
  loc_00DF9EB2: push ecx
  loc_00DF9EB3: push edx
  loc_00DF9EB4: call 006DD568h ; SelectObject(%x1v, %x2v)
  loc_00DF9EB9: call edi
  loc_00DF9EBB: push ebx
  loc_00DF9EBC: call 006DD498h ; DeleteObject(%x1v)
  loc_00DF9EC1: call edi
  loc_00DF9EC3: pop edi
  loc_00DF9EC4: pop esi
  loc_00DF9EC5: pop ebp
  loc_00DF9EC6: xor eax, eax
  loc_00DF9EC8: pop ebx
  loc_00DF9EC9: pop ecx
  loc_00DF9ECA: retn 000Ch
End Sub

Private Sub Proc_2_137_DF9ED0
  loc_00DF9ED0: push ebp
  loc_00DF9ED1: mov ebp, esp
  loc_00DF9ED3: sub esp, 00000008h
  loc_00DF9ED6: push 00402836h ; __vbaExceptHandler
  loc_00DF9EDB: mov eax, fs:[00000000h]
  loc_00DF9EE1: push eax
  loc_00DF9EE2: mov fs:[00000000h], esp
  loc_00DF9EE9: sub esp, 00000078h
  loc_00DF9EEC: push ebx
  loc_00DF9EED: push esi
  loc_00DF9EEE: push edi
  loc_00DF9EEF: mov var_8, esp
  loc_00DF9EF2: mov var_4, 00401878h
  loc_00DF9EF9: mov esi, Me
  loc_00DF9EFC: mov ebx, [004010B4h] ; __vbaObjSetAddref
  loc_00DF9F02: xor eax, eax
  loc_00DF9F04: lea edx, var_24
  loc_00DF9F07: mov ecx, [esi+000000E8h]
  loc_00DF9F0D: xor edi, edi
  loc_00DF9F0F: mov var_2C, eax
  loc_00DF9F12: push ecx
  loc_00DF9F13: push edx
  loc_00DF9F14: mov var_14, edi
  loc_00DF9F17: mov var_20, edi
  loc_00DF9F1A: mov var_24, edi
  loc_00DF9F1D: mov var_28, eax
  loc_00DF9F20: mov var_30, edi
  loc_00DF9F23: mov var_34, edi
  loc_00DF9F26: mov var_44, edi
  loc_00DF9F29: mov var_54, edi
  loc_00DF9F2C: mov var_64, edi
  loc_00DF9F2F: mov var_74, edi
  loc_00DF9F32: mov var_78, edi
  loc_00DF9F35: mov var_7C, edi
  loc_00DF9F38: call ebx
  loc_00DF9F3A: mov ecx, var_14
  loc_00DF9F3D: mov eax, [esi+000000ECh]
  loc_00DF9F43: lea edx, var_14
  loc_00DF9F46: push ecx
  loc_00DF9F47: push edx
  loc_00DF9F48: mov var_18, eax
  loc_00DF9F4B: call ebx
  loc_00DF9F4D: mov eax, [esi]
  loc_00DF9F4F: lea ecx, var_78
  loc_00DF9F52: push ecx
  loc_00DF9F53: push esi
  loc_00DF9F54: call [eax+00000108h]
  loc_00DF9F5A: cmp eax, edi
  loc_00DF9F5C: fnclex
  loc_00DF9F5E: jge 00DF9F72h
  loc_00DF9F60: push 00000108h
  loc_00DF9F65: push 006DCB00h
  loc_00DF9F6A: push esi
  loc_00DF9F6B: push eax
  loc_00DF9F6C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF9F72: fld real4 ptr var_78
  loc_00DF9F75: mov ebx, [00401208h] ; __vbaFpI4
  loc_00DF9F7B: call ebx
  loc_00DF9F7D: mov edx, [esi]
  loc_00DF9F7F: mov [esi+000001D0h], eax
  loc_00DF9F85: lea eax, var_78
  loc_00DF9F88: push eax
  loc_00DF9F89: push esi
  loc_00DF9F8A: call [edx+00000100h]
  loc_00DF9F90: cmp eax, edi
  loc_00DF9F92: fnclex
  loc_00DF9F94: jge 00DF9FA8h
  loc_00DF9F96: push 00000100h
  loc_00DF9F9B: push 006DCB00h
  loc_00DF9FA0: push esi
  loc_00DF9FA1: push eax
  loc_00DF9FA2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DF9FA8: fld real4 ptr var_78
  loc_00DF9FAB: call ebx
  loc_00DF9FAD: mov ecx, var_18
  loc_00DF9FB0: mov [esi+000001D4h], eax
  loc_00DF9FB6: cmp ecx, 00000007h
  loc_00DF9FB9: mov [esi+000000E4h], FFFFFFh
  loc_00DF9FC2: ja 00DFA03Fh
  loc_00DF9FC4: jmp [ecx*4+00DFA2A0h]
  loc_00DF9FCB: mov ecx, [esi+000001D0h]
  loc_00DF9FD1: mov var_20, ecx
  loc_00DF9FD4: jmp 00DFA046h
  loc_00DF9FD6: mov ecx, [esi+000000F0h]
  loc_00DF9FDC: mov eax, [esi+000000ECh]
  loc_00DF9FE2: or ecx, 00000008h
  loc_00DF9FE5: cmp eax, 00000005h
  loc_00DF9FE8: mov [esi+000000F0h], ecx
  loc_00DF9FEE: jnz 00DFA046h
  loc_00DF9FF0: mov edx, [esi+000001D0h]
  loc_00DF9FF6: mov var_20, edx
  loc_00DF9FF9: jmp 00DFA046h
  loc_00DF9FFB: mov var_30, eax
  loc_00DF9FFE: mov eax, [esi+000000ECh]
  loc_00DFA004: cmp eax, 00000007h
  loc_00DFA007: jnz 00DFA046h
  loc_00DFA009: mov eax, [esi+000001D0h]
  loc_00DFA00F: mov var_20, eax
  loc_00DFA012: jmp 00DFA046h
  loc_00DFA014: mov ecx, [esi+000000ECh]
  loc_00DFA01A: mov [esi+000000F0h], 00000020h
  loc_00DFA024: cmp ecx, 00000006h
  loc_00DFA027: jnz 00DFA02Eh
  loc_00DFA029: mov var_30, eax
  loc_00DFA02C: jmp 00DFA046h
  loc_00DFA02E: cmp ecx, 00000004h
  loc_00DFA031: jnz 00DFA046h
  loc_00DFA033: mov [esi+000000F0h], 00000028h
  loc_00DFA03D: jmp 00DFA046h
  loc_00DFA03F: mov [esi+000000E4h], di
  loc_00DFA046: cmp [esi+000000E4h], di
  loc_00DFA04D: jz 00DFA266h
  loc_00DFA053: mov ecx, var_14
  loc_00DFA056: push edi
  loc_00DFA057: push ecx
  loc_00DFA058: call [0040114Ch] ; __vbaObjIs
  loc_00DFA05E: test ax, ax
  loc_00DFA061: jz 00DFA117h
  loc_00DFA067: sub esp, 00000010h
  loc_00DFA06A: mov ecx, 0000000Ah
  loc_00DFA06F: mov ebx, esp
  loc_00DFA071: mov eax, 80020004h
  loc_00DFA076: sub esp, 00000010h
  loc_00DFA079: mov edx, 00000003h
  loc_00DFA07E: mov [ebx], ecx
  loc_00DFA080: mov ecx, var_70
  loc_00DFA083: mov var_44, edx
  loc_00DFA086: mov edi, [esi+00000010h]
  loc_00DFA089: mov [ebx+00000004h], ecx
  loc_00DFA08C: mov ecx, esp
  loc_00DFA08E: sub esp, 00000010h
  loc_00DFA091: mov edi, [edi]
  loc_00DFA093: mov [ebx+00000008h], eax
  loc_00DFA096: mov eax, var_68
  loc_00DFA099: mov [ebx+0000000Ch], eax
  loc_00DFA09C: mov eax, var_20
  loc_00DFA09F: mov [ecx], edx
  loc_00DFA0A1: mov edx, var_60
  loc_00DFA0A4: mov [ecx+00000004h], edx
  loc_00DFA0A7: mov edx, var_58
  loc_00DFA0AA: mov [ecx+00000008h], eax
  loc_00DFA0AD: mov eax, 00000003h
  loc_00DFA0B2: mov [ecx+0000000Ch], edx
  loc_00DFA0B5: mov edx, var_50
  loc_00DFA0B8: mov ecx, esp
  loc_00DFA0BA: sub esp, 00000010h
  loc_00DFA0BD: mov [ecx], eax
  loc_00DFA0BF: mov eax, var_30
  loc_00DFA0C2: mov [ecx+00000004h], edx
  loc_00DFA0C5: mov edx, var_48
  loc_00DFA0C8: mov [ecx+00000008h], eax
  loc_00DFA0CB: mov eax, var_40
  loc_00DFA0CE: mov [ecx+0000000Ch], edx
  loc_00DFA0D1: mov edx, var_44
  loc_00DFA0D4: mov ecx, esp
  loc_00DFA0D6: mov [ecx], edx
  loc_00DFA0D8: mov edx, var_38
  loc_00DFA0DB: mov [ecx+00000004h], eax
  loc_00DFA0DE: mov eax, [esi+000000F0h]
  loc_00DFA0E4: mov [ecx+00000008h], eax
  loc_00DFA0E7: mov eax, [esi+000000E8h]
  loc_00DFA0ED: push eax
  loc_00DFA0EE: mov [ecx+0000000Ch], edx
  loc_00DFA0F1: lea ecx, var_34
  loc_00DFA0F4: push ecx
  loc_00DFA0F5: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFA0FB: mov edx, [esi+00000010h]
  loc_00DFA0FE: push eax
  loc_00DFA0FF: push edx
  loc_00DFA100: call [edi+00000354h]
  loc_00DFA106: xor ebx, ebx
  loc_00DFA108: cmp eax, ebx
  loc_00DFA10A: fnclex
  loc_00DFA10C: jge 00DFA1D1h
  loc_00DFA112: jmp 00DFA1BCh
  loc_00DFA117: sub esp, 00000010h
  loc_00DFA11A: mov eax, var_14
  loc_00DFA11D: mov ebx, esp
  loc_00DFA11F: mov ecx, 00000009h
  loc_00DFA124: sub esp, 00000010h
  loc_00DFA127: mov edx, 00000003h
  loc_00DFA12C: mov [ebx], ecx
  loc_00DFA12E: mov ecx, var_70
  loc_00DFA131: mov var_44, edx
  loc_00DFA134: mov edi, [esi+00000010h]
  loc_00DFA137: mov [ebx+00000004h], ecx
  loc_00DFA13A: mov ecx, esp
  loc_00DFA13C: sub esp, 00000010h
  loc_00DFA13F: mov edi, [edi]
  loc_00DFA141: mov [ebx+00000008h], eax
  loc_00DFA144: mov eax, var_68
  loc_00DFA147: mov [ebx+0000000Ch], eax
  loc_00DFA14A: mov eax, var_20
  loc_00DFA14D: mov [ecx], edx
  loc_00DFA14F: mov edx, var_60
  loc_00DFA152: mov [ecx+00000004h], edx
  loc_00DFA155: mov edx, var_58
  loc_00DFA158: mov [ecx+00000008h], eax
  loc_00DFA15B: mov eax, 00000003h
  loc_00DFA160: mov [ecx+0000000Ch], edx
  loc_00DFA163: mov edx, var_50
  loc_00DFA166: mov ecx, esp
  loc_00DFA168: sub esp, 00000010h
  loc_00DFA16B: mov [ecx], eax
  loc_00DFA16D: mov eax, var_30
  loc_00DFA170: mov [ecx+00000004h], edx
  loc_00DFA173: mov edx, var_48
  loc_00DFA176: mov [ecx+00000008h], eax
  loc_00DFA179: mov eax, var_40
  loc_00DFA17C: mov [ecx+0000000Ch], edx
  loc_00DFA17F: mov edx, var_44
  loc_00DFA182: mov ecx, esp
  loc_00DFA184: mov [ecx], edx
  loc_00DFA186: mov edx, var_38
  loc_00DFA189: mov [ecx+00000004h], eax
  loc_00DFA18C: mov eax, [esi+000000F0h]
  loc_00DFA192: mov [ecx+00000008h], eax
  loc_00DFA195: mov eax, [esi+000000E8h]
  loc_00DFA19B: push eax
  loc_00DFA19C: mov [ecx+0000000Ch], edx
  loc_00DFA19F: lea ecx, var_34
  loc_00DFA1A2: push ecx
  loc_00DFA1A3: call [004010B4h] ; __vbaObjSetAddref
  loc_00DFA1A9: mov edx, [esi+00000010h]
  loc_00DFA1AC: push eax
  loc_00DFA1AD: push edx
  loc_00DFA1AE: call [edi+00000354h]
  loc_00DFA1B4: xor ebx, ebx
  loc_00DFA1B6: cmp eax, ebx
  loc_00DFA1B8: fnclex
  loc_00DFA1BA: jge 00DFA1D1h
  loc_00DFA1BC: mov ecx, [esi+00000010h]
  loc_00DFA1BF: push 00000354h
  loc_00DFA1C4: push 006DCB00h
  loc_00DFA1C9: push ecx
  loc_00DFA1CA: push eax
  loc_00DFA1CB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFA1D1: lea ecx, var_34
  loc_00DFA1D4: call [00401254h] ; __vbaFreeObj
  loc_00DFA1DA: lea edx, var_2C
  loc_00DFA1DD: push edx
  loc_00DFA1DE: call 006DDC7Ch ; GetCursorPos(%x1v)
  loc_00DFA1E3: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DFA1E9: call edi
  loc_00DFA1EB: mov eax, var_28
  loc_00DFA1EE: mov ecx, var_2C
  loc_00DFA1F1: push eax
  loc_00DFA1F2: push ecx
  loc_00DFA1F3: call 006DDC34h ; ChildWindowFromPoint(%x1v, %x2v)
  loc_00DFA1F8: mov var_78, eax
  loc_00DFA1FB: call edi
  loc_00DFA1FD: mov eax, [esi+00000010h]
  loc_00DFA200: lea ecx, var_7C
  loc_00DFA203: push ecx
  loc_00DFA204: push eax
  loc_00DFA205: mov edx, [eax]
  loc_00DFA207: call [edx+00000058h]
  loc_00DFA20A: cmp eax, ebx
  loc_00DFA20C: fnclex
  loc_00DFA20E: jge 00DFA222h
  loc_00DFA210: mov edx, [esi+00000010h]
  loc_00DFA213: push 00000058h
  loc_00DFA215: push 006DCB00h
  loc_00DFA21A: push edx
  loc_00DFA21B: push eax
  loc_00DFA21C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFA222: mov eax, var_78
  loc_00DFA225: mov ecx, var_7C
  loc_00DFA228: cmp eax, ecx
  loc_00DFA22A: jnz 00DFA23Dh
  loc_00DFA22C: mov [esi+000000E2h], FFFFFFh
  loc_00DFA235: push 00DFA289h
  loc_00DFA23A: fwait
  loc_00DFA23B: jmp 00DFA278h
  loc_00DFA23D: mov ecx, [esi]
  loc_00DFA23F: push esi
  loc_00DFA240: mov [esi+0000004Ch], bx
  loc_00DFA244: mov [esi+0000004Eh], bx
  loc_00DFA248: mov [esi+000000A8h], bx
  loc_00DFA24F: mov [esi+00000048h], ebx
  loc_00DFA252: mov [esi+000000E2h], bx
  loc_00DFA259: mov [esi+000000E4h], bx
  loc_00DFA260: call [ecx+00000910h]
  loc_00DFA266: fwait
  loc_00DFA267: push 00DFA289h
  loc_00DFA26C: jmp 00DFA278h
  loc_00DFA26E: lea ecx, var_34
  loc_00DFA271: call [00401254h] ; __vbaFreeObj
  loc_00DFA277: ret
  loc_00DFA278: mov esi, [00401254h] ; __vbaFreeObj
  loc_00DFA27E: lea ecx, var_14
  loc_00DFA281: call __vbaFreeObj
  loc_00DFA283: lea ecx, var_24
  loc_00DFA286: call __vbaFreeObj
  loc_00DFA288: ret
  loc_00DFA289: mov ecx, var_10
  loc_00DFA28C: pop edi
  loc_00DFA28D: pop esi
  loc_00DFA28E: xor eax, eax
  loc_00DFA290: mov fs:[00000000h], ecx
  loc_00DFA297: pop ebx
  loc_00DFA298: mov esp, ebp
  loc_00DFA29A: pop ebp
  loc_00DFA29B: retn 0004h
End Sub

Private Sub Proc_2_138_DFA2C0(arg_C, arg_10) 'DFA2C0
  loc_00DFA2C0: sub esp, 00000010h
  loc_00DFA2C3: mov eax, arg_10
  loc_00DFA2C7: push ebx
  loc_00DFA2C8: push ebp
  loc_00DFA2C9: push esi
  loc_00DFA2CA: mov ecx, [eax]
  loc_00DFA2CC: push edi
  loc_00DFA2CD: mov edx, ecx
  loc_00DFA2CF: mov eax, ecx
  loc_00DFA2D1: and edx, 000000FFh
  loc_00DFA2D7: mov edi, arg_24
  loc_00DFA2DB: mov arg_20, edx
  loc_00DFA2DF: mov ebx, [00401208h] ; __vbaFpI4
  loc_00DFA2E5: cdq
  loc_00DFA2E6: fld real4 ptr [edi]
  loc_00DFA2E8: and edx, 000000FFh
  loc_00DFA2EE: add eax, edx
  loc_00DFA2F0: sar eax, 08h
  loc_00DFA2F3: and eax, 000000FFh
  loc_00DFA2F8: mov Me, eax
  loc_00DFA2FC: mov eax, ecx
  loc_00DFA2FE: fmul st0, real4 ptr [0040188Ch]
  loc_00DFA304: cdq
  loc_00DFA305: and edx, 0000FFFFh
  loc_00DFA30B: add eax, edx
  loc_00DFA30D: sar eax, 10h
  loc_00DFA310: and eax, 000000FFh
  loc_00DFA315: mov arg_C, eax
  loc_00DFA319: fnstsw ax
  loc_00DFA31B: test al, 0Dh
  loc_00DFA31D: jnz 00DFA405h
  loc_00DFA323: fild real4 ptr arg_20
  loc_00DFA327: fstp real8 ptr arg_10
  loc_00DFA32B: fadd st0, real8 ptr arg_10
  loc_00DFA32F: fnstsw ax
  loc_00DFA331: test al, 0Dh
  loc_00DFA333: jnz 00DFA405h
  loc_00DFA339: call ebx
  loc_00DFA33B: fld real4 ptr [edi]
  loc_00DFA33D: fmul st0, real4 ptr [0040188Ch]
  loc_00DFA343: mov ebp, eax
  loc_00DFA345: fnstsw ax
  loc_00DFA347: test al, 0Dh
  loc_00DFA349: jnz 00DFA405h
  loc_00DFA34F: fild real4 ptr Me
  loc_00DFA353: fstp real8 ptr arg_10
  loc_00DFA357: fadd st0, real8 ptr arg_10
  loc_00DFA35B: fnstsw ax
  loc_00DFA35D: test al, 0Dh
  loc_00DFA35F: jnz 00DFA405h
  loc_00DFA365: call ebx
  loc_00DFA367: fld real4 ptr [edi]
  loc_00DFA369: fmul st0, real4 ptr [0040188Ch]
  loc_00DFA36F: mov esi, eax
  loc_00DFA371: fnstsw ax
  loc_00DFA373: test al, 0Dh
  loc_00DFA375: jnz 00DFA405h
  loc_00DFA37B: fild real4 ptr arg_C
  loc_00DFA37F: fstp real8 ptr arg_10
  loc_00DFA383: fadd st0, real8 ptr arg_10
  loc_00DFA387: fnstsw ax
  loc_00DFA389: test al, 0Dh
  loc_00DFA38B: jnz 00DFA405h
  loc_00DFA38D: call ebx
  loc_00DFA38F: fld real4 ptr [edi]
  loc_00DFA391: fcomp real4 ptr [00401888h]
  loc_00DFA397: mov ecx, eax
  loc_00DFA399: fnstsw ax
  loc_00DFA39B: test ah, 41h
  loc_00DFA39E: jnz 00DFA3C9h
  loc_00DFA3A0: cmp ebp, 000000FFh
  loc_00DFA3A6: jle 00DFA3ADh
  loc_00DFA3A8: mov ebp, 000000FFh
  loc_00DFA3AD: cmp esi, 000000FFh
  loc_00DFA3B3: jle 00DFA3BAh
  loc_00DFA3B5: mov esi, 000000FFh
  loc_00DFA3BA: cmp ecx, 000000FFh
  loc_00DFA3C0: jle 00DFA3DBh
  loc_00DFA3C2: mov ecx, 000000FFh
  loc_00DFA3C7: jmp 00DFA3DBh
  loc_00DFA3C9: test ebp, ebp
  loc_00DFA3CB: jge 00DFA3CFh
  loc_00DFA3CD: xor ebp, ebp
  loc_00DFA3CF: test esi, esi
  loc_00DFA3D1: jge 00DFA3D5h
  loc_00DFA3D3: xor esi, esi
  loc_00DFA3D5: test ecx, ecx
  loc_00DFA3D7: jge 00DFA3DBh
  loc_00DFA3D9: xor ecx, ecx
  loc_00DFA3DB: imul esi, esi, 00000100h
  loc_00DFA3E1: jo 00DFA40Ah
  loc_00DFA3E3: add esi, ebp
  loc_00DFA3E5: mov eax, arg_28
  loc_00DFA3E9: jo 00DFA40Ah
  loc_00DFA3EB: imul ecx, ecx, 00010000h
  loc_00DFA3F1: jo 00DFA40Ah
  loc_00DFA3F3: add esi, ecx
  loc_00DFA3F5: pop edi
  loc_00DFA3F6: jo 00DFA40Ah
  loc_00DFA3F8: mov [eax], esi
  loc_00DFA3FA: pop esi
  loc_00DFA3FB: pop ebp
  loc_00DFA3FC: xor eax, eax
  loc_00DFA3FE: pop ebx
  loc_00DFA3FF: add esp, 00000010h
  loc_00DFA402: retn 0010h
End Sub

Private Sub Proc_2_139_DFB480
  loc_00DFB480: push ebp
  loc_00DFB481: mov ebp, esp
  loc_00DFB483: sub esp, 00000008h
  loc_00DFB486: push 00402836h ; __vbaExceptHandler
  loc_00DFB48B: mov eax, fs:[00000000h]
  loc_00DFB491: push eax
  loc_00DFB492: mov fs:[00000000h], esp
  loc_00DFB499: sub esp, 00000108h
  loc_00DFB49F: push ebx
  loc_00DFB4A0: push esi
  loc_00DFB4A1: push edi
  loc_00DFB4A2: mov var_8, esp
  loc_00DFB4A5: mov var_4, 00401918h
  loc_00DFB4AC: mov ecx, 0000000Bh
  loc_00DFB4B1: xor eax, eax
  loc_00DFB4B3: lea edi, var_3C
  loc_00DFB4B6: mov esi, Me
  loc_00DFB4B9: repz stosd
  loc_00DFB4BB: mov ecx, 0000000Bh
  loc_00DFB4C0: lea edi, var_6C
  loc_00DFB4C3: repz stosd
  loc_00DFB4C5: mov var_7C, eax
  loc_00DFB4C8: xor ebx, ebx
  loc_00DFB4CA: cmp [esi+0000007Ch], bx
  loc_00DFB4CE: mov var_78, eax
  loc_00DFB4D1: mov var_74, eax
  loc_00DFB4D4: mov ecx, 0000000Bh
  loc_00DFB4D9: lea edi, var_114
  loc_00DFB4DF: mov var_70, eax
  loc_00DFB4E2: mov var_84, ebx
  loc_00DFB4E8: mov var_88, ebx
  loc_00DFB4EE: mov var_98, ebx
  loc_00DFB4F4: mov var_A8, ebx
  loc_00DFB4FA: mov var_B8, ebx
  loc_00DFB500: mov var_BC, ebx
  loc_00DFB506: mov var_C0, ebx
  loc_00DFB50C: mov var_C4, ebx
  loc_00DFB512: mov var_C8, ebx
  loc_00DFB518: repz stosd
  loc_00DFB51A: jz 00DFBBD6h
  loc_00DFB520: cmp [esi+000000E2h], bx
  loc_00DFB527: jnz 00DFBBD6h
  loc_00DFB52D: cmp [esi+00000048h], 00000002h
  loc_00DFB531: jz 00DFBBD6h
  loc_00DFB537: mov eax, [esi+00000110h]
  loc_00DFB53D: cmp eax, ebx
  loc_00DFB53F: jz 00DFB54Dh
  loc_00DFB541: push eax
  loc_00DFB542: call 006DDEE8h ; DestroyWindow(%x1v)
  loc_00DFB547: call [00401074h] ; __vbaSetSystemError
  loc_00DFB54D: mov eax, [esi+00000104h]
  loc_00DFB553: mov var_40, 00000003h
  loc_00DFB55A: cmp eax, 00000001h
  loc_00DFB55D: jnz 00DFB566h
  loc_00DFB55F: mov var_40, 00000043h
  loc_00DFB566: cmp [esi+0000010Eh], bx
  loc_00DFB56D: mov eax, [esi+00000010h]
  loc_00DFB570: jz 00DFB666h
  loc_00DFB576: mov ecx, [eax]
  loc_00DFB578: lea edx, var_C0
  loc_00DFB57E: push edx
  loc_00DFB57F: push eax
  loc_00DFB580: call [ecx+00000058h]
  loc_00DFB583: cmp eax, ebx
  loc_00DFB585: fnclex
  loc_00DFB587: jge 00DFB59Fh
  loc_00DFB589: mov ecx, [esi+00000010h]
  loc_00DFB58C: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFB592: push 00000058h
  loc_00DFB594: push 006DCB00h
  loc_00DFB599: push ecx
  loc_00DFB59A: push eax
  loc_00DFB59B: call ebx
  loc_00DFB59D: jmp 00DFB5A5h
  loc_00DFB59F: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFB5A5: mov eax, [00E3D9CCh]
  loc_00DFB5AA: test eax, eax
  loc_00DFB5AC: jnz 00DFB5BEh
  loc_00DFB5AE: push 00E3D9CCh
  loc_00DFB5B3: push 006DC960h
  loc_00DFB5B8: call [004011A0h] ; __vbaNew2
  loc_00DFB5BE: mov edi, [00E3D9CCh]
  loc_00DFB5C4: lea eax, var_88
  loc_00DFB5CA: push eax
  loc_00DFB5CB: push edi
  loc_00DFB5CC: mov edx, [edi]
  loc_00DFB5CE: call [edx+00000014h]
  loc_00DFB5D1: test eax, eax
  loc_00DFB5D3: fnclex
  loc_00DFB5D5: jge 00DFB5E2h
  loc_00DFB5D7: push 00000014h
  loc_00DFB5D9: push 006DC950h
  loc_00DFB5DE: push edi
  loc_00DFB5DF: push eax
  loc_00DFB5E0: call ebx
  loc_00DFB5E2: mov eax, var_88
  loc_00DFB5E8: lea edx, var_C4
  loc_00DFB5EE: push edx
  loc_00DFB5EF: push eax
  loc_00DFB5F0: mov ecx, [eax]
  loc_00DFB5F2: mov edi, eax
  loc_00DFB5F4: call [ecx+00000100h]
  loc_00DFB5FA: test eax, eax
  loc_00DFB5FC: fnclex
  loc_00DFB5FE: jge 00DFB60Eh
  loc_00DFB600: push 00000100h
  loc_00DFB605: push 006DEB4Ch
  loc_00DFB60A: push edi
  loc_00DFB60B: push eax
  loc_00DFB60C: call ebx
  loc_00DFB60E: mov ecx, var_C4
  loc_00DFB614: mov edx, var_C0
  loc_00DFB61A: lea eax, var_C8
  loc_00DFB620: mov var_C8, 00000000h
  loc_00DFB62A: push eax
  loc_00DFB62B: mov eax, var_40
  loc_00DFB62E: push ecx
  loc_00DFB62F: push 00000000h
  loc_00DFB631: push edx
  loc_00DFB632: push 80000000h
  loc_00DFB637: push 80000000h
  loc_00DFB63C: push 80000000h
  loc_00DFB641: push 80000000h
  loc_00DFB646: push eax
  loc_00DFB647: push 00000000h
  loc_00DFB649: lea ecx, var_84
  loc_00DFB64F: push 006DC5E0h ; "tooltips_class32"
  loc_00DFB654: push ecx
  loc_00DFB655: call [004011ECh] ; __vbaStrToAnsi
  loc_00DFB65B: push eax
  loc_00DFB65C: push 00400000h
  loc_00DFB661: jmp 00DFB74Eh
  loc_00DFB666: mov edx, [eax]
  loc_00DFB668: lea ecx, var_C0
  loc_00DFB66E: push ecx
  loc_00DFB66F: push eax
  loc_00DFB670: call [edx+00000058h]
  loc_00DFB673: cmp eax, ebx
  loc_00DFB675: fnclex
  loc_00DFB677: jge 00DFB68Fh
  loc_00DFB679: mov edx, [esi+00000010h]
  loc_00DFB67C: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFB682: push 00000058h
  loc_00DFB684: push 006DCB00h
  loc_00DFB689: push edx
  loc_00DFB68A: push eax
  loc_00DFB68B: call ebx
  loc_00DFB68D: jmp 00DFB695h
  loc_00DFB68F: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DFB695: mov eax, [00E3D9CCh]
  loc_00DFB69A: test eax, eax
  loc_00DFB69C: jnz 00DFB6AEh
  loc_00DFB69E: push 00E3D9CCh
  loc_00DFB6A3: push 006DC960h
  loc_00DFB6A8: call [004011A0h] ; __vbaNew2
  loc_00DFB6AE: mov edi, [00E3D9CCh]
  loc_00DFB6B4: lea ecx, var_88
  loc_00DFB6BA: push ecx
  loc_00DFB6BB: push edi
  loc_00DFB6BC: mov eax, [edi]
  loc_00DFB6BE: call [eax+00000014h]
  loc_00DFB6C1: test eax, eax
  loc_00DFB6C3: fnclex
  loc_00DFB6C5: jge 00DFB6D2h
  loc_00DFB6C7: push 00000014h
  loc_00DFB6C9: push 006DC950h
  loc_00DFB6CE: push edi
  loc_00DFB6CF: push eax
  loc_00DFB6D0: call ebx
  loc_00DFB6D2: mov eax, var_88
  loc_00DFB6D8: lea ecx, var_C4
  loc_00DFB6DE: push ecx
  loc_00DFB6DF: push eax
  loc_00DFB6E0: mov edx, [eax]
  loc_00DFB6E2: mov edi, eax
  loc_00DFB6E4: call [edx+00000100h]
  loc_00DFB6EA: test eax, eax
  loc_00DFB6EC: fnclex
  loc_00DFB6EE: jge 00DFB6FEh
  loc_00DFB6F0: push 00000100h
  loc_00DFB6F5: push 006DEB4Ch
  loc_00DFB6FA: push edi
  loc_00DFB6FB: push eax
  loc_00DFB6FC: call ebx
  loc_00DFB6FE: mov eax, var_C4
  loc_00DFB704: mov ecx, var_C0
  loc_00DFB70A: lea edx, var_C8
  loc_00DFB710: mov var_C8, 00000000h
  loc_00DFB71A: push edx
  loc_00DFB71B: mov edx, var_40
  loc_00DFB71E: push eax
  loc_00DFB71F: push 00000000h
  loc_00DFB721: push ecx
  loc_00DFB722: push 80000000h
  loc_00DFB727: push 80000000h
  loc_00DFB72C: push 80000000h
  loc_00DFB731: push 80000000h
  loc_00DFB736: push edx
  loc_00DFB737: push 00000000h
  loc_00DFB739: lea eax, var_84
  loc_00DFB73F: push 006DC5E0h ; "tooltips_class32"
  loc_00DFB744: push eax
  loc_00DFB745: call [004011ECh] ; __vbaStrToAnsi
  loc_00DFB74B: push eax
  loc_00DFB74C: push 00000000h
  loc_00DFB74E: call 006DDE54h ; CreateWindowEx(%x1v, %x2v, %x3v, %x4v, %x5v, %x6v, %x7v, %x8v, %x9v, %x10v, %x11v, %x12v)
  loc_00DFB753: mov edi, eax
  loc_00DFB755: call [00401074h] ; __vbaSetSystemError
  loc_00DFB75B: lea ecx, var_84
  loc_00DFB761: mov [esi+00000110h], edi
  loc_00DFB767: call [00401258h] ; __vbaFreeStr
  loc_00DFB76D: lea ecx, var_88
  loc_00DFB773: call [00401254h] ; __vbaFreeObj
  loc_00DFB779: mov ecx, [esi+00000110h]
  loc_00DFB77F: push FFFFFFE6h
  loc_00DFB781: push ecx
  loc_00DFB782: call 006DDF30h ; GetClassLong(%x1v, %x2v)
  loc_00DFB787: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00DFB78D: mov var_C0, eax
  loc_00DFB793: call edi
  loc_00DFB795: mov edx, var_C0
  loc_00DFB79B: mov eax, [esi+00000110h]
  loc_00DFB7A1: or edx, 00020000h
  loc_00DFB7A7: push edx
  loc_00DFB7A8: push FFFFFFE6h
  loc_00DFB7AA: push eax
  loc_00DFB7AB: call 006DDFC0h ; SetClassLong(%x1v, %x2v, %x3v)
  loc_00DFB7B0: call edi
  loc_00DFB7B2: mov eax, [esi+00000010h]
  loc_00DFB7B5: lea edx, var_C0
  loc_00DFB7BB: push edx
  loc_00DFB7BC: push eax
  loc_00DFB7BD: mov ecx, [eax]
  loc_00DFB7BF: call [ecx+00000058h]
  loc_00DFB7C2: test eax, eax
  loc_00DFB7C4: fnclex
  loc_00DFB7C6: jge 00DFB7D6h
  loc_00DFB7C8: mov ecx, [esi+00000010h]
  loc_00DFB7CB: push 00000058h
  loc_00DFB7CD: push 006DCB00h
  loc_00DFB7D2: push ecx
  loc_00DFB7D3: push eax
  loc_00DFB7D4: call ebx
  loc_00DFB7D6: mov eax, var_C0
  loc_00DFB7DC: lea edx, var_7C
  loc_00DFB7DF: push edx
  loc_00DFB7E0: push eax
  loc_00DFB7E1: call 006DDAA8h ; GetClientRect(%x1v, %x2v)
  loc_00DFB7E6: call edi
  loc_00DFB7E8: mov eax, [esi+00000078h]
  loc_00DFB7EB: test eax, eax
  loc_00DFB7ED: jz 00DFB929h
  loc_00DFB7F3: mov cx, [esi+0000010Ch]
  loc_00DFB7FA: mov eax, [esi+00000010h]
  loc_00DFB7FD: neg cx
  loc_00DFB800: sbb ecx, ecx
  loc_00DFB802: and ecx, 00000002h
  loc_00DFB805: add ecx, 00000011h
  loc_00DFB808: mov var_68, ecx
  loc_00DFB80B: mov edx, [eax]
  loc_00DFB80D: lea ecx, var_C0
  loc_00DFB813: push ecx
  loc_00DFB814: push eax
  loc_00DFB815: call [edx+00000058h]
  loc_00DFB818: test eax, eax
  loc_00DFB81A: fnclex
  loc_00DFB81C: jge 00DFB82Ch
  loc_00DFB81E: mov edx, [esi+00000010h]
  loc_00DFB821: push 00000058h
  loc_00DFB823: push 006DCB00h
  loc_00DFB828: push edx
  loc_00DFB829: push eax
  loc_00DFB82A: call ebx
  loc_00DFB82C: mov eax, var_C0
  loc_00DFB832: mov ecx, [esi]
  loc_00DFB834: lea edx, var_C0
  loc_00DFB83A: mov var_64, eax
  loc_00DFB83D: push edx
  loc_00DFB83E: push esi
  loc_00DFB83F: call [ecx+00000818h]
  loc_00DFB845: test eax, eax
  loc_00DFB847: jge 00DFB857h
  loc_00DFB849: push 00000818h
  loc_00DFB84E: push 006DD090h
  loc_00DFB853: push esi
  loc_00DFB854: push eax
  loc_00DFB855: call ebx
  loc_00DFB857: mov eax, var_C0
  loc_00DFB85D: mov var_6C, 0000002Ch
  loc_00DFB864: mov var_60, eax
  loc_00DFB867: mov eax, [00E3D9CCh]
  loc_00DFB86C: test eax, eax
  loc_00DFB86E: jnz 00DFB880h
  loc_00DFB870: push 00E3D9CCh
  loc_00DFB875: push 006DC960h
  loc_00DFB87A: call [004011A0h] ; __vbaNew2
  loc_00DFB880: mov edi, [00E3D9CCh]
  loc_00DFB886: lea edx, var_88
  loc_00DFB88C: push edx
  loc_00DFB88D: push edi
  loc_00DFB88E: mov ecx, [edi]
  loc_00DFB890: call [ecx+00000014h]
  loc_00DFB893: test eax, eax
  loc_00DFB895: fnclex
  loc_00DFB897: jge 00DFB8A4h
  loc_00DFB899: push 00000014h
  loc_00DFB89B: push 006DC950h
  loc_00DFB8A0: push edi
  loc_00DFB8A1: push eax
  loc_00DFB8A2: call ebx
  loc_00DFB8A4: mov eax, var_88
  loc_00DFB8AA: lea edx, var_C0
  loc_00DFB8B0: push edx
  loc_00DFB8B1: push eax
  loc_00DFB8B2: mov ecx, [eax]
  loc_00DFB8B4: mov edi, eax
  loc_00DFB8B6: call [ecx+00000100h]
  loc_00DFB8BC: test eax, eax
  loc_00DFB8BE: fnclex
  loc_00DFB8C0: jge 00DFB8D0h
  loc_00DFB8C2: push 00000100h
  loc_00DFB8C7: push 006DEB4Ch
  loc_00DFB8CC: push edi
  loc_00DFB8CD: push eax
  loc_00DFB8CE: call ebx
  loc_00DFB8D0: mov eax, var_C0
  loc_00DFB8D6: lea ecx, var_88
  loc_00DFB8DC: mov var_4C, eax
  loc_00DFB8DF: call [00401254h] ; __vbaFreeObj
  loc_00DFB8E5: mov ecx, [esi+000000F8h]
  loc_00DFB8EB: push ecx
  loc_00DFB8EC: call [00401190h] ; VarPtr
  loc_00DFB8F2: mov var_48, eax
  loc_00DFB8F5: lea edx, var_7C
  loc_00DFB8F8: lea eax, var_5C
  loc_00DFB8FB: push edx
  loc_00DFB8FC: push eax
  loc_00DFB8FD: push 00000010h
  loc_00DFB8FF: call [00401060h] ; __vbaCopyBytes
  loc_00DFB905: mov edx, [esi+00000110h]
  loc_00DFB90B: lea ecx, var_6C
  loc_00DFB90E: push ecx
  loc_00DFB90F: push 00000000h
  loc_00DFB911: push 00000432h
  loc_00DFB916: push edx
  loc_00DFB917: call 006DDEA0h ; SendMessage(%x1v, %x2v, %x3v, %x4v)
  loc_00DFB91C: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00DFB922: call ebx
  loc_00DFB924: jmp 00DFBA98h
  loc_00DFB929: mov ax, [esi+0000010Ch]
  loc_00DFB930: lea edx, var_C0
  loc_00DFB936: neg ax
  loc_00DFB939: sbb eax, eax
  loc_00DFB93B: push edx
  loc_00DFB93C: and eax, 00000002h
  loc_00DFB93F: add eax, 00000010h
  loc_00DFB942: mov var_38, eax
  loc_00DFB945: mov eax, [esi+00000010h]
  loc_00DFB948: push eax
  loc_00DFB949: mov ecx, [eax]
  loc_00DFB94B: call [ecx+00000058h]
  loc_00DFB94E: test eax, eax
  loc_00DFB950: fnclex
  loc_00DFB952: jge 00DFB962h
  loc_00DFB954: mov ecx, [esi+00000010h]
  loc_00DFB957: push 00000058h
  loc_00DFB959: push 006DCB00h
  loc_00DFB95E: push ecx
  loc_00DFB95F: push eax
  loc_00DFB960: call ebx
  loc_00DFB962: mov edx, var_C0
  loc_00DFB968: mov eax, [esi]
  loc_00DFB96A: lea ecx, var_C0
  loc_00DFB970: mov var_34, edx
  loc_00DFB973: push ecx
  loc_00DFB974: push esi
  loc_00DFB975: call [eax+00000818h]
  loc_00DFB97B: test eax, eax
  loc_00DFB97D: jge 00DFB98Dh
  loc_00DFB97F: push 00000818h
  loc_00DFB984: push 006DD090h
  loc_00DFB989: push esi
  loc_00DFB98A: push eax
  loc_00DFB98B: call ebx
  loc_00DFB98D: mov eax, [00E3D9CCh]
  loc_00DFB992: mov edx, var_C0
  loc_00DFB998: test eax, eax
  loc_00DFB99A: mov var_30, edx
  loc_00DFB99D: mov var_3C, 0000002Ch
  loc_00DFB9A4: jnz 00DFB9B6h
  loc_00DFB9A6: push 00E3D9CCh
  loc_00DFB9AB: push 006DC960h
  loc_00DFB9B0: call [004011A0h] ; __vbaNew2
  loc_00DFB9B6: mov edi, [00E3D9CCh]
  loc_00DFB9BC: lea ecx, var_88
  loc_00DFB9C2: push ecx
  loc_00DFB9C3: push edi
  loc_00DFB9C4: mov eax, [edi]
  loc_00DFB9C6: call [eax+00000014h]
  loc_00DFB9C9: test eax, eax
  loc_00DFB9CB: fnclex
  loc_00DFB9CD: jge 00DFB9DAh
  loc_00DFB9CF: push 00000014h
  loc_00DFB9D1: push 006DC950h
  loc_00DFB9D6: push edi
  loc_00DFB9D7: push eax
  loc_00DFB9D8: call ebx
  loc_00DFB9DA: mov eax, var_88
  loc_00DFB9E0: lea ecx, var_C0
  loc_00DFB9E6: push ecx
  loc_00DFB9E7: push eax
  loc_00DFB9E8: mov edx, [eax]
  loc_00DFB9EA: mov edi, eax
  loc_00DFB9EC: call [edx+00000100h]
  loc_00DFB9F2: test eax, eax
  loc_00DFB9F4: fnclex
  loc_00DFB9F6: jge 00DFBA06h
  loc_00DFB9F8: push 00000100h
  loc_00DFB9FD: push 006DEB4Ch
  loc_00DFBA02: push edi
  loc_00DFBA03: push eax
  loc_00DFBA04: call ebx
  loc_00DFBA06: mov edx, var_C0
  loc_00DFBA0C: lea ecx, var_88
  loc_00DFBA12: mov var_1C, edx
  loc_00DFBA15: call [00401254h] ; __vbaFreeObj
  loc_00DFBA1B: mov edx, [esi+000000F8h]
  loc_00DFBA21: lea ecx, var_18
  loc_00DFBA24: call [004011B0h] ; __vbaStrCopy
  loc_00DFBA2A: lea eax, var_7C
  loc_00DFBA2D: lea ecx, var_2C
  loc_00DFBA30: push eax
  loc_00DFBA31: push ecx
  loc_00DFBA32: push 00000010h
  loc_00DFBA34: call [00401060h] ; __vbaCopyBytes
  loc_00DFBA3A: lea edx, var_3C
  loc_00DFBA3D: lea eax, var_114
  loc_00DFBA43: push edx
  loc_00DFBA44: push eax
  loc_00DFBA45: push 006DD194h ; UDT_5_006DD194
  loc_00DFBA4A: call [0040113Ch] ; __vbaRecUniToAnsi
  loc_00DFBA50: mov ecx, [esi+00000110h]
  loc_00DFBA56: push eax
  loc_00DFBA57: push 00000000h
  loc_00DFBA59: push 00000404h
  loc_00DFBA5E: push ecx
  loc_00DFBA5F: call 006DDEA0h ; SendMessage(%x1v, %x2v, %x3v, %x4v)
  loc_00DFBA64: call [00401074h] ; __vbaSetSystemError
  loc_00DFBA6A: lea edx, var_114
  loc_00DFBA70: lea eax, var_3C
  loc_00DFBA73: push edx
  loc_00DFBA74: push eax
  loc_00DFBA75: push 006DD194h ; UDT_5_006DD194
  loc_00DFBA7A: call [00401058h] ; __vbaRecAnsiToUni
  loc_00DFBA80: lea ecx, var_114
  loc_00DFBA86: push ecx
  loc_00DFBA87: push 006DD194h ; UDT_5_006DD194
  loc_00DFBA8C: call [00401218h] ; __vbaRecDestructAnsi
  loc_00DFBA92: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00DFBA98: mov edx, [esi+000000FCh]
  loc_00DFBA9E: lea edi, [esi+000000FCh]
  loc_00DFBAA4: push edx
  loc_00DFBAA5: push 00000000h
  loc_00DFBAA7: call [00401104h] ; __vbaStrCmp
  loc_00DFBAAD: test eax, eax
  loc_00DFBAAF: jnz 00DFBABBh
  loc_00DFBAB1: mov eax, [esi+00000100h]
  loc_00DFBAB7: test eax, eax
  loc_00DFBAB9: jz 00DFBB2Eh
  loc_00DFBABB: mov eax, [esi+00000078h]
  loc_00DFBABE: test eax, eax
  loc_00DFBAC0: mov eax, [edi]
  loc_00DFBAC2: push eax
  loc_00DFBAC3: jz 00DFBAECh
  loc_00DFBAC5: call [00401190h] ; VarPtr
  loc_00DFBACB: test eax, eax
  loc_00DFBACD: jz 00DFBB2Eh
  loc_00DFBACF: mov ecx, [esi+00000100h]
  loc_00DFBAD5: mov edx, [esi+00000110h]
  loc_00DFBADB: push eax
  loc_00DFBADC: push ecx
  loc_00DFBADD: push 00000421h
  loc_00DFBAE2: push edx
  loc_00DFBAE3: call 006DDEA0h ; SendMessage(%x1v, %x2v, %x3v, %x4v)
  loc_00DFBAE8: call ebx
  loc_00DFBAEA: jmp 00DFBB2Eh
  loc_00DFBAEC: lea ecx, var_84
  loc_00DFBAF2: push ecx
  loc_00DFBAF3: call [004011ECh] ; __vbaStrToAnsi
  loc_00DFBAF9: mov edx, [esi+00000100h]
  loc_00DFBAFF: push eax
  loc_00DFBB00: mov eax, [esi+00000110h]
  loc_00DFBB06: push edx
  loc_00DFBB07: push 00000420h
  loc_00DFBB0C: push eax
  loc_00DFBB0D: call 006DDEA0h ; SendMessage(%x1v, %x2v, %x3v, %x4v)
  loc_00DFBB12: call ebx
  loc_00DFBB14: mov ecx, var_84
  loc_00DFBB1A: push ecx
  loc_00DFBB1B: push edi
  loc_00DFBB1C: call [00401160h] ; __vbaStrToUnicode
  loc_00DFBB22: lea ecx, var_84
  loc_00DFBB28: call [00401258h] ; __vbaFreeStr
  loc_00DFBB2E: mov eax, [esi+00000110h]
  loc_00DFBB34: lea edx, var_BC
  loc_00DFBB3A: push edx
  loc_00DFBB3B: push 00000000h
  loc_00DFBB3D: push 00000418h
  loc_00DFBB42: push eax
  loc_00DFBB43: mov var_BC, 000000F0h
  loc_00DFBB4D: call 006DDEA0h ; SendMessage(%x1v, %x2v, %x3v, %x4v)
  loc_00DFBB52: call ebx
  loc_00DFBB54: mov ecx, [esi+00000108h]
  loc_00DFBB5A: lea edx, var_B8
  loc_00DFBB60: lea eax, var_A8
  loc_00DFBB66: xor edi, edi
  loc_00DFBB68: push edx
  loc_00DFBB69: push eax
  loc_00DFBB6A: mov var_B0, ecx
  loc_00DFBB70: mov var_B8, 00008003h
  loc_00DFBB7A: mov var_A8, edi
  loc_00DFBB80: call [004011CCh] ; __vbaVarTstNe
  loc_00DFBB86: test ax, ax
  loc_00DFBB89: jz 00DFBBD6h
  loc_00DFBB8B: mov ecx, [esi]
  loc_00DFBB8D: lea edx, var_C4
  loc_00DFBB93: push edx
  loc_00DFBB94: mov edx, [esi+00000108h]
  loc_00DFBB9A: lea eax, var_C0
  loc_00DFBBA0: mov var_C0, edi
  loc_00DFBBA6: push eax
  loc_00DFBBA7: push edx
  loc_00DFBBA8: push esi
  loc_00DFBBA9: call [ecx+0000090Ch]
  loc_00DFBBAF: mov ecx, var_C4
  loc_00DFBBB5: mov edx, [esi+00000110h]
  loc_00DFBBBB: lea eax, var_C8
  loc_00DFBBC1: mov var_C8, edi
  loc_00DFBBC7: push eax
  loc_00DFBBC8: push ecx
  loc_00DFBBC9: push 00000413h
  loc_00DFBBCE: push edx
  loc_00DFBBCF: call 006DDEA0h ; SendMessage(%x1v, %x2v, %x3v, %x4v)
  loc_00DFBBD4: call ebx
  loc_00DFBBD6: push 00DFBC24h
  loc_00DFBBDB: jmp 00DFBC02h
  loc_00DFBBDD: lea ecx, var_84
  loc_00DFBBE3: call [00401258h] ; __vbaFreeStr
  loc_00DFBBE9: lea ecx, var_88
  loc_00DFBBEF: call [00401254h] ; __vbaFreeObj
  loc_00DFBBF5: lea ecx, var_98
  loc_00DFBBFB: call [00401024h] ; __vbaFreeVar
  loc_00DFBC01: ret
  loc_00DFBC02: lea eax, var_114
  loc_00DFBC08: push eax
  loc_00DFBC09: push 006DD194h ; UDT_5_006DD194
  loc_00DFBC0E: call [00401218h] ; __vbaRecDestructAnsi
  loc_00DFBC14: lea ecx, var_3C
  loc_00DFBC17: push ecx
  loc_00DFBC18: push 006DD194h ; UDT_5_006DD194
  loc_00DFBC1D: call [00401078h] ; __vbaRecDestruct
  loc_00DFBC23: ret
  loc_00DFBC24: mov ecx, var_10
  loc_00DFBC27: pop edi
  loc_00DFBC28: pop esi
  loc_00DFBC29: xor eax, eax
  loc_00DFBC2B: mov fs:[00000000h], ecx
  loc_00DFBC32: pop ebx
  loc_00DFBC33: mov esp, ebp
  loc_00DFBC35: pop ebp
  loc_00DFBC36: retn 0004h
End Sub

Private Sub Proc_2_140_DFBC40
  loc_00DFBC40: mov ecx, var_4
  loc_00DFBC44: mov eax, [ecx+00000044h]
  loc_00DFBC47: cmp eax, 0000000Ah
  loc_00DFBC4A: ja 00DFBC6Ch
  loc_00DFBC4C: jmp [eax*4+00DFBC74h]
  loc_00DFBC53: mov [ecx+000000CCh], 00000000h
  loc_00DFBC5D: xor eax, eax
  loc_00DFBC5F: retn 0004h
End Sub

Private Sub Proc_2_141_DFBCA0(arg_C, arg_10, arg_14, arg_18) 'DFBCA0
  loc_00DFBCA0: sub esp, 00000014h
  loc_00DFBCA3: push esi
  loc_00DFBCA4: mov esi, arg_14
  loc_00DFBCA8: push edi
  loc_00DFBCA9: xor edi, edi
  loc_00DFBCAB: mov eax, [esi+00000044h]
  loc_00DFBCAE: mov [esp+00000008h], edi
  loc_00DFBCB2: cmp eax, 0000000Ch
  loc_00DFBCB5: mov Proc_2_141_DFBCA0, edi
  loc_00DFBCB9: mov arg_C, edi
  loc_00DFBCBD: mov Me, edi
  loc_00DFBCC1: mov arg_10, edi
  loc_00DFBCC5: ja 00DFBFA4h
  loc_00DFBCCB: jmp [eax*4+00DFC000h]
  loc_00DFBCD2: push 0000000Fh
  loc_00DFBCD4: call 006DDDB4h ; GetSysColor(%x1v)
  loc_00DFBCD9: mov [esp+00000008h], eax
  loc_00DFBCDD: call [00401074h] ; __vbaSetSystemError
  loc_00DFBCE3: mov eax, [esp+00000008h]
  loc_00DFBCE7: mov [esi+00000088h], eax
  loc_00DFBCED: jmp 00DFBFA4h
  loc_00DFBCF2: mov eax, [esi+000000CCh]
  loc_00DFBCF8: sub eax, edi
  loc_00DFBCFA: jz 00DFBD3Ah
  loc_00DFBCFC: dec eax
  loc_00DFBCFD: jz 00DFBD20h
  loc_00DFBCFF: dec eax
  loc_00DFBD00: jnz 00DFBFA4h
  loc_00DFBD06: mov ecx, [esi]
  loc_00DFBD08: lea edx, Proc_2_141_DFBCA0
  loc_00DFBD0C: lea eax, [esp+00000008h]
  loc_00DFBD10: push edx
  loc_00DFBD11: push eax
  loc_00DFBD12: mov Me, edi
  loc_00DFBD16: push 00FCF1F0h
  loc_00DFBD1B: jmp 00DFBF93h
  loc_00DFBD20: mov edx, [esi]
  loc_00DFBD22: lea eax, Proc_2_141_DFBCA0
  loc_00DFBD26: lea ecx, [esp+00000008h]
  loc_00DFBD2A: push eax
  loc_00DFBD2B: push ecx
  loc_00DFBD2C: mov Me, edi
  loc_00DFBD30: push 00DBEEF3h
  loc_00DFBD35: jmp 00DFBEC6h
  loc_00DFBD3A: mov eax, [esi]
  loc_00DFBD3C: lea ecx, Proc_2_141_DFBCA0
  loc_00DFBD40: lea edx, [esp+00000008h]
  loc_00DFBD44: push ecx
  loc_00DFBD45: push edx
  loc_00DFBD46: mov Me, edi
  loc_00DFBD4A: push 00E7EBECh
  loc_00DFBD4F: jmp 00DFBEF1h
  loc_00DFBD54: mov eax, [esi+000000CCh]
  loc_00DFBD5A: sub eax, edi
  loc_00DFBD5C: jz 00DFBDB4h
  loc_00DFBD5E: dec eax
  loc_00DFBD5F: jz 00DFBD8Ch
  loc_00DFBD61: dec eax
  loc_00DFBD62: jnz 00DFBDDAh
  loc_00DFBD64: mov ecx, [esi]
  loc_00DFBD66: lea edx, Proc_2_141_DFBCA0
  loc_00DFBD6A: lea eax, [esp+00000008h]
  loc_00DFBD6E: push edx
  loc_00DFBD6F: push eax
  loc_00DFBD70: push 00E3DFE0h
  loc_00DFBD75: push esi
  loc_00DFBD76: mov arg_10, edi
  loc_00DFBD7A: call [ecx+0000090Ch]
  loc_00DFBD80: mov ecx, Proc_2_141_DFBCA0
  loc_00DFBD84: mov [esi+00000088h], ecx
  loc_00DFBD8A: jmp 00DFBDDAh
  loc_00DFBD8C: mov edx, [esi]
  loc_00DFBD8E: lea eax, Proc_2_141_DFBCA0
  loc_00DFBD92: lea ecx, [esp+00000008h]
  loc_00DFBD96: push eax
  loc_00DFBD97: push ecx
  loc_00DFBD98: push 00BAD6D4h
  loc_00DFBD9D: push esi
  loc_00DFBD9E: mov arg_10, edi
  loc_00DFBDA2: call [edx+0000090Ch]
  loc_00DFBDA8: mov edx, Proc_2_141_DFBCA0
  loc_00DFBDAC: mov [esi+00000088h], edx
  loc_00DFBDB2: jmp 00DFBDDAh
  loc_00DFBDB4: mov eax, [esi]
  loc_00DFBDB6: lea ecx, Proc_2_141_DFBCA0
  loc_00DFBDBA: lea edx, [esp+00000008h]
  loc_00DFBDBE: push ecx
  loc_00DFBDBF: push edx
  loc_00DFBDC0: push 00FFD1ADh
  loc_00DFBDC5: push esi
  loc_00DFBDC6: mov arg_10, edi
  loc_00DFBDCA: call [eax+0000090Ch]
  loc_00DFBDD0: mov eax, Proc_2_141_DFBCA0
  loc_00DFBDD4: mov [esi+00000088h], eax
  loc_00DFBDDA: mov ecx, [esi]
  loc_00DFBDDC: lea edx, Proc_2_141_DFBCA0
  loc_00DFBDE0: lea eax, [esp+00000008h]
  loc_00DFBDE4: push edx
  loc_00DFBDE5: push eax
  loc_00DFBDE6: push 008B4215h
  loc_00DFBDEB: push esi
  loc_00DFBDEC: mov arg_10, edi
  loc_00DFBDF0: call [ecx+0000090Ch]
  loc_00DFBDF6: mov ecx, Proc_2_141_DFBCA0
  loc_00DFBDFA: mov [esi+00000090h], ecx
  loc_00DFBE00: jmp 00DFBFA4h
  loc_00DFBE05: mov eax, [esi+000000CCh]
  loc_00DFBE0B: sub eax, edi
  loc_00DFBE0D: jz 00DFBE83h
  loc_00DFBE0F: dec eax
  loc_00DFBE10: jz 00DFBE6Ch
  loc_00DFBE12: dec eax
  loc_00DFBE13: jnz 00DFBFA4h
  loc_00DFBE19: mov edx, [esi]
  loc_00DFBE1B: lea eax, Proc_2_141_DFBCA0
  loc_00DFBE1F: lea ecx, [esp+00000008h]
  loc_00DFBE23: push eax
  loc_00DFBE24: push ecx
  loc_00DFBE25: push 00D4D4D4h
  loc_00DFBE2A: push esi
  loc_00DFBE2B: mov arg_10, edi
  loc_00DFBE2F: call [edx+0000090Ch]
  loc_00DFBE35: mov eax, [esi]
  loc_00DFBE37: lea ecx, arg_10
  loc_00DFBE3B: mov edx, Proc_2_141_DFBCA0
  loc_00DFBE3F: push ecx
  loc_00DFBE40: mov arg_10, edx
  loc_00DFBE44: lea edx, arg_C
  loc_00DFBE48: lea ecx, arg_10
  loc_00DFBE4C: push edx
  loc_00DFBE4D: push ecx
  loc_00DFBE4E: push esi
  loc_00DFBE4F: mov arg_18, 3D75C28Fh
  loc_00DFBE57: call [eax+0000097Ch]
  loc_00DFBE5D: mov edx, arg_10
  loc_00DFBE61: mov [esi+00000088h], edx
  loc_00DFBE67: jmp 00DFBFA4h
  loc_00DFBE6C: mov eax, [esi]
  loc_00DFBE6E: lea ecx, Proc_2_141_DFBCA0
  loc_00DFBE72: lea edx, [esp+00000008h]
  loc_00DFBE76: push ecx
  loc_00DFBE77: push edx
  loc_00DFBE78: mov Me, edi
  loc_00DFBE7C: push 00DEEDE8h
  loc_00DFBE81: jmp 00DFBEF1h
  loc_00DFBE83: mov ecx, [esi]
  loc_00DFBE85: lea edx, Proc_2_141_DFBCA0
  loc_00DFBE89: lea eax, [esp+00000008h]
  loc_00DFBE8D: push edx
  loc_00DFBE8E: push eax
  loc_00DFBE8F: mov Me, edi
  loc_00DFBE93: push 00FDECE0h
  loc_00DFBE98: jmp 00DFBF93h
  loc_00DFBE9D: mov eax, [esi+000000CCh]
  loc_00DFBEA3: sub eax, edi
  loc_00DFBEA5: jz 00DFBF07h
  loc_00DFBEA7: dec eax
  loc_00DFBEA8: jz 00DFBEDCh
  loc_00DFBEAA: dec eax
  loc_00DFBEAB: jnz 00DFBFA4h
  loc_00DFBEB1: mov edx, [esi]
  loc_00DFBEB3: lea eax, Proc_2_141_DFBCA0
  loc_00DFBEB7: lea ecx, [esp+00000008h]
  loc_00DFBEBB: push eax
  loc_00DFBEBC: push ecx
  loc_00DFBEBD: mov Me, edi
  loc_00DFBEC1: push 00E1D6D5h
  loc_00DFBEC6: push esi
  loc_00DFBEC7: call [edx+0000090Ch]
  loc_00DFBECD: mov edx, Proc_2_141_DFBCA0
  loc_00DFBED1: mov [esi+00000088h], edx
  loc_00DFBED7: jmp 00DFBFA4h
  loc_00DFBEDC: mov eax, [esi]
  loc_00DFBEDE: lea ecx, Proc_2_141_DFBCA0
  loc_00DFBEE2: lea edx, [esp+00000008h]
  loc_00DFBEE6: push ecx
  loc_00DFBEE7: push edx
  loc_00DFBEE8: mov Me, edi
  loc_00DFBEEC: push 00BAD6D4h
  loc_00DFBEF1: push esi
  loc_00DFBEF2: call [eax+0000090Ch]
  loc_00DFBEF8: mov eax, Proc_2_141_DFBCA0
  loc_00DFBEFC: mov [esi+00000088h], eax
  loc_00DFBF02: jmp 00DFBFA4h
  loc_00DFBF07: mov ecx, [esi]
  loc_00DFBF09: lea edx, Proc_2_141_DFBCA0
  loc_00DFBF0D: lea eax, [esp+00000008h]
  loc_00DFBF11: push edx
  loc_00DFBF12: push eax
  loc_00DFBF13: mov Me, edi
  loc_00DFBF17: push 00FFD1ADh
  loc_00DFBF1C: jmp 00DFBF93h
  loc_00DFBF1E: mov eax, [esi+000000CCh]
  loc_00DFBF24: sub eax, edi
  loc_00DFBF26: jz 00DFBF7Eh
  loc_00DFBF28: dec eax
  loc_00DFBF29: jz 00DFBEDCh
  loc_00DFBF2B: dec eax
  loc_00DFBF2C: jnz 00DFBFA4h
  loc_00DFBF2E: mov edx, [esi]
  loc_00DFBF30: lea eax, Proc_2_141_DFBCA0
  loc_00DFBF34: lea ecx, [esp+00000008h]
  loc_00DFBF38: push eax
  loc_00DFBF39: push ecx
  loc_00DFBF3A: push 00BA9EA0h
  loc_00DFBF3F: push esi
  loc_00DFBF40: mov arg_10, edi
  loc_00DFBF44: call [edx+0000090Ch]
  loc_00DFBF4A: mov eax, [esi]
  loc_00DFBF4C: lea ecx, arg_10
  loc_00DFBF50: mov edx, Proc_2_141_DFBCA0
  loc_00DFBF54: push ecx
  loc_00DFBF55: mov arg_10, edx
  loc_00DFBF59: lea edx, arg_C
  loc_00DFBF5D: lea ecx, arg_10
  loc_00DFBF61: push edx
  loc_00DFBF62: push ecx
  loc_00DFBF63: push esi
  loc_00DFBF64: mov arg_18, 3E19999Ah
  loc_00DFBF6C: call [eax+0000097Ch]
  loc_00DFBF72: mov edx, arg_10
  loc_00DFBF76: mov [esi+00000088h], edx
  loc_00DFBF7C: jmp 00DFBFA4h
  loc_00DFBF7E: mov ecx, [esi]
  loc_00DFBF80: lea edx, Proc_2_141_DFBCA0
  loc_00DFBF84: lea eax, [esp+00000008h]
  loc_00DFBF88: push edx
  loc_00DFBF89: push eax
  loc_00DFBF8A: mov Me, edi
  loc_00DFBF8E: push 00FCE1CAh
  loc_00DFBF93: push esi
  loc_00DFBF94: call [ecx+0000090Ch]
  loc_00DFBF9A: mov ecx, Proc_2_141_DFBCA0
  loc_00DFBF9E: mov [esi+00000088h], ecx
  loc_00DFBFA4: mov edx, [esi]
  loc_00DFBFA6: lea eax, Proc_2_141_DFBCA0
  loc_00DFBFAA: lea ecx, [esp+00000008h]
  loc_00DFBFAE: push eax
  loc_00DFBFAF: push ecx
  loc_00DFBFB0: push 80000012h
  loc_00DFBFB5: push esi
  loc_00DFBFB6: mov arg_10, edi
  loc_00DFBFBA: call [edx+0000090Ch]
  loc_00DFBFC0: or ecx, FFFFFFFFh
  loc_00DFBFC3: mov edx, Proc_2_141_DFBCA0
  loc_00DFBFC7: mov [esi+00000090h], edx
  loc_00DFBFCD: mov eax, [esi+00000044h]
  loc_00DFBFD0: cmp eax, 00000001h
  loc_00DFBFD3: jz 00DFBFE4h
  loc_00DFBFD5: cmp eax, 00000009h
  loc_00DFBFD8: jz 00DFBFE4h
  loc_00DFBFDA: cmp eax, edi
  loc_00DFBFDC: jz 00DFBFE4h
  loc_00DFBFDE: mov [esi+0000006Eh], di
  loc_00DFBFE2: jmp 00DFBFE8h
  loc_00DFBFE4: mov [esi+0000006Eh], cx
  loc_00DFBFE8: cmp eax, 00000004h
  loc_00DFBFEB: jnz 00DFBFF4h
  loc_00DFBFED: mov [esi+00000170h], cx
  loc_00DFBFF4: pop edi
  loc_00DFBFF5: xor eax, eax
  loc_00DFBFF7: pop esi
  loc_00DFBFF8: add esp, 00000014h
  loc_00DFBFFB: retn 0004h
End Sub

Private Sub Proc_2_142_E00000(arg_C, arg_10) 'E00000
  loc_00E00000: push ebp
  loc_00E00001: mov ebp, esp
  loc_00E00003: sub esp, 00000008h
  loc_00E00006: push 00402836h ; __vbaExceptHandler
  loc_00E0000B: mov eax, fs:[00000000h]
  loc_00E00011: push eax
  loc_00E00012: mov fs:[00000000h], esp
  loc_00E00019: sub esp, 00000054h
  loc_00E0001C: push ebx
  loc_00E0001D: push esi
  loc_00E0001E: push edi
  loc_00E0001F: mov var_8, esp
  loc_00E00022: mov var_4, 00401B00h
  loc_00E00029: mov ecx, 00000006h
  loc_00E0002E: xor eax, eax
  loc_00E00030: lea edi, var_28
  loc_00E00033: mov ebx, [00401214h] ; __vbaLateMemCallLd
  loc_00E00039: repz stosd
  loc_00E0003B: mov edi, arg_C
  loc_00E0003E: mov var_2C, eax
  loc_00E00041: mov var_3C, eax
  loc_00E00044: mov var_4C, eax
  loc_00E00047: push eax
  loc_00E00048: mov eax, [edi]
  loc_00E0004A: push 006DF09Ch ; "Type"
  loc_00E0004F: lea ecx, var_3C
  loc_00E00052: push eax
  loc_00E00053: push ecx
  loc_00E00054: mov var_54, 00000001h
  loc_00E0005B: mov var_5C, 00008003h
  loc_00E00062: call ebx
  loc_00E00064: add esp, 00000010h
  loc_00E00067: lea edx, var_5C
  loc_00E0006A: push eax
  loc_00E0006B: push edx
  loc_00E0006C: call [00401108h] ; __vbaVarTstEq
  loc_00E00072: lea ecx, var_3C
  loc_00E00075: mov esi, eax
  loc_00E00077: call [00401024h] ; __vbaFreeVar
  loc_00E0007D: test si, si
  loc_00E00080: jz 00E000C9h
  loc_00E00082: mov eax, [edi]
  loc_00E00084: push 00000000h
  loc_00E00086: push 006DF0A8h ; "Handle"
  loc_00E0008B: lea ecx, var_3C
  loc_00E0008E: push eax
  loc_00E0008F: push ecx
  loc_00E00090: call ebx
  loc_00E00092: add esp, 00000010h
  loc_00E00095: lea edx, var_28
  loc_00E00098: lea eax, var_3C
  loc_00E0009B: push edx
  loc_00E0009C: push 00000018h
  loc_00E0009E: push eax
  loc_00E0009F: call [004011D0h] ; __vbaI4Var
  loc_00E000A5: push eax
  loc_00E000A6: call 006DD8F8h ; GetObject(%x1v, %x2v, %x3v)
  loc_00E000AB: call [00401074h] ; __vbaSetSystemError
  loc_00E000B1: lea ecx, var_3C
  loc_00E000B4: call [00401024h] ; __vbaFreeVar
  loc_00E000BA: xor ecx, ecx
  loc_00E000BC: cmp var_16, 0020h
  loc_00E000C1: setz cl
  loc_00E000C4: neg ecx
  loc_00E000C6: mov var_2C, ecx
  loc_00E000C9: push 00E000E5h
  loc_00E000CE: jmp 00E000E4h
  loc_00E000D0: lea edx, var_4C
  loc_00E000D3: lea eax, var_3C
  loc_00E000D6: push edx
  loc_00E000D7: push eax
  loc_00E000D8: push 00000002h
  loc_00E000DA: call [00401038h] ; __vbaFreeVarList
  loc_00E000E0: add esp, 0000000Ch
  loc_00E000E3: ret
  loc_00E000E4: ret
  loc_00E000E5: mov ecx, arg_10
  loc_00E000E8: mov dx, var_2C
  loc_00E000EC: pop edi
  loc_00E000ED: pop esi
  loc_00E000EE: mov [ecx], dx
  loc_00E000F1: mov ecx, var_10
  loc_00E000F4: xor eax, eax
  loc_00E000F6: mov fs:[00000000h], ecx
  loc_00E000FD: pop ebx
  loc_00E000FE: mov esp, ebp
  loc_00E00100: pop ebp
  loc_00E00101: retn 000Ch
End Sub

Private Sub Proc_2_143_E00110(arg_C, arg_10) 'E00110
  loc_00E00110: push ebp
  loc_00E00111: mov ebp, esp
  loc_00E00113: sub esp, 00000014h
  loc_00E00116: push 00402836h ; __vbaExceptHandler
  loc_00E0011B: mov eax, fs:[00000000h]
  loc_00E00121: push eax
  loc_00E00122: mov fs:[00000000h], esp
  loc_00E00129: sub esp, 00000024h
  loc_00E0012C: push ebx
  loc_00E0012D: push esi
  loc_00E0012E: push edi
  loc_00E0012F: mov var_14, esp
  loc_00E00132: mov var_10, 00401B10h
  loc_00E00139: xor ebx, ebx
  loc_00E0013B: mov var_C, ebx
  loc_00E0013E: mov var_8, ebx
  loc_00E00141: mov var_20, ebx
  loc_00E00144: mov var_24, ebx
  loc_00E00147: mov var_2C, ebx
  loc_00E0014A: mov edx, arg_C
  loc_00E0014D: lea ecx, var_20
  loc_00E00150: call [004011B0h] ; __vbaStrCopy
  loc_00E00156: push 00000001h
  loc_00E00158: call [004010A4h] ; __vbaOnError
  loc_00E0015E: mov eax, var_20
  loc_00E00161: push eax
  loc_00E00162: lea ecx, var_2C
  loc_00E00165: push ecx
  loc_00E00166: call [004011ECh] ; __vbaStrToAnsi
  loc_00E0016C: push eax
  loc_00E0016D: call 006DE4D0h ; LoadLibrary(%x1v)
  loc_00E00172: mov esi, eax
  loc_00E00174: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00E0017A: call edi
  loc_00E0017C: mov edx, var_2C
  loc_00E0017F: push edx
  loc_00E00180: lea eax, var_20
  loc_00E00183: push eax
  loc_00E00184: call [00401160h] ; __vbaStrToUnicode
  loc_00E0018A: lea ecx, var_2C
  loc_00E0018D: call [00401258h] ; __vbaFreeStr
  loc_00E00193: cmp esi, ebx
  loc_00E00195: jz 00E001A6h
  loc_00E00197: push esi
  loc_00E00198: call 006DE488h ; FreeLibrary(%x1v)
  loc_00E0019D: call edi
  loc_00E0019F: mov var_24, FFFFFFFFh
  loc_00E001A6: call [00401098h] ; __vbaExitProc
  loc_00E001AC: push 00E001C7h
  loc_00E001B1: jmp 00E001BDh
  loc_00E001B3: lea ecx, var_2C
  loc_00E001B6: call [00401258h] ; __vbaFreeStr
  loc_00E001BC: ret
  loc_00E001BD: lea ecx, var_20
  loc_00E001C0: call [00401258h] ; __vbaFreeStr
  loc_00E001C6: ret
  loc_00E001C7: mov ecx, arg_10
  loc_00E001CA: mov dx, var_24
  loc_00E001CE: mov [ecx], dx
  loc_00E001D1: xor eax, eax
  loc_00E001D3: mov ecx, var_1C
  loc_00E001D6: mov fs:[00000000h], ecx
  loc_00E001DD: pop edi
  loc_00E001DE: pop esi
  loc_00E001DF: pop ebx
  loc_00E001E0: mov esp, ebp
  loc_00E001E2: pop ebp
  loc_00E001E3: retn 000Ch
End Sub

Private Sub Proc_2_144_E001F0(arg_C) 'E001F0
  loc_00E001F0: push ebp
  loc_00E001F1: mov ebp, esp
  loc_00E001F3: sub esp, 00000018h
  loc_00E001F6: push 00402836h ; __vbaExceptHandler
  loc_00E001FB: mov eax, fs:[00000000h]
  loc_00E00201: push eax
  loc_00E00202: mov fs:[00000000h], esp
  loc_00E00209: mov eax, 00000040h
  loc_00E0020E: call 00402830h ; __vbaChkstk
  loc_00E00213: push ebx
  loc_00E00214: push esi
  loc_00E00215: push edi
  loc_00E00216: mov var_18, esp
  loc_00E00219: mov var_14, 00401B38h ; "$"
  loc_00E00220: mov var_10, 00000000h
  loc_00E00227: mov var_C, 00000000h
  loc_00E0022E: mov var_4, 00000001h
  loc_00E00235: mov var_4, 00000002h
  loc_00E0023C: push FFFFFFFFh
  loc_00E0023E: call [004010A4h] ; __vbaOnError
  loc_00E00244: mov var_4, 00000003h
  loc_00E0024B: lea eax, var_48
  loc_00E0024E: push eax
  loc_00E0024F: mov ecx, Me
  loc_00E00252: mov edx, [ecx]
  loc_00E00254: mov eax, Me
  loc_00E00257: push eax
  loc_00E00258: call [edx+000009E8h]
  loc_00E0025E: movsx ecx, var_48
  loc_00E00262: test ecx, ecx
  loc_00E00264: jz 00E002E0h
  loc_00E00266: mov var_4, 00000004h
  loc_00E0026D: mov edx, Me
  loc_00E00270: mov eax, [edx+000001E8h]
  loc_00E00276: push eax
  loc_00E00277: lea ecx, var_34
  loc_00E0027A: push ecx
  loc_00E0027B: call [004011B8h] ; __vbaVarNot
  loc_00E00281: push eax
  loc_00E00282: call [004010DCh] ; __vbaBoolVarNull
  loc_00E00288: movsx edx, ax
  loc_00E0028B: test edx, edx
  loc_00E0028D: jz 00E002E0h
  loc_00E0028F: mov var_4, 00000005h
  loc_00E00296: call 006DE14Ch ; IsAppThemed()
  loc_00E0029B: mov var_4C, eax
  loc_00E0029E: call [00401074h] ; __vbaSetSystemError
  loc_00E002A4: xor eax, eax
  loc_00E002A6: cmp var_4C, 00000000h
  loc_00E002AA: setnz al
  loc_00E002AD: neg eax
  loc_00E002AF: mov ecx, Me
  loc_00E002B2: mov [ecx+000000D0h], ax
  loc_00E002B9: mov var_4, 00000006h
  loc_00E002C0: mov var_3C, FFFFFFFFh
  loc_00E002C7: mov var_44, 0000000Bh
  loc_00E002CE: lea edx, var_44
  loc_00E002D1: mov eax, Me
  loc_00E002D4: mov ecx, [eax+000001E8h]
  loc_00E002DA: call [0040101Ch] ; __vbaVarMove
  loc_00E002E0: mov var_4, 00000009h
  loc_00E002E7: mov ecx, Me
  loc_00E002EA: mov dx, [ecx+000000D0h]
  loc_00E002F1: mov var_24, dx
  loc_00E002F5: push 00E00307h
  loc_00E002FA: jmp 00E00306h
  loc_00E002FC: lea ecx, var_34
  loc_00E002FF: call [00401024h] ; __vbaFreeVar
  loc_00E00305: ret
  loc_00E00306: ret
  loc_00E00307: mov eax, arg_C
  loc_00E0030A: mov cx, var_24
  loc_00E0030E: mov [eax], cx
  loc_00E00311: xor eax, eax
  loc_00E00313: mov ecx, var_20
  loc_00E00316: mov fs:[00000000h], ecx
  loc_00E0031D: pop edi
  loc_00E0031E: pop esi
  loc_00E0031F: pop ebx
  loc_00E00320: mov esp, ebp
  loc_00E00322: pop ebp
  loc_00E00323: retn 0008h
End Sub

Private Sub Proc_2_145_E00330(arg_C) 'E00330
  loc_00E00330: push ebp
  loc_00E00331: mov ebp, esp
  loc_00E00333: sub esp, 00000008h
  loc_00E00336: push 00402836h ; __vbaExceptHandler
  loc_00E0033B: mov eax, fs:[00000000h]
  loc_00E00341: push eax
  loc_00E00342: mov fs:[00000000h], esp
  loc_00E00349: sub esp, 0000002Ch
  loc_00E0034C: push ebx
  loc_00E0034D: push esi
  loc_00E0034E: push edi
  loc_00E0034F: mov var_8, esp
  loc_00E00352: mov var_4, 00401B80h
  loc_00E00359: mov esi, Me
  loc_00E0035C: xor eax, eax
  loc_00E0035E: mov var_24, eax
  loc_00E00361: mov var_34, eax
  loc_00E00364: mov var_38, eax
  loc_00E00367: mov eax, [esi+000001E8h]
  loc_00E0036D: add eax, 00000010h
  loc_00E00370: lea ecx, var_24
  loc_00E00373: push eax
  loc_00E00374: push ecx
  loc_00E00375: call [004011B8h] ; __vbaVarNot
  loc_00E0037B: push eax
  loc_00E0037C: call [004010DCh] ; __vbaBoolVarNull
  loc_00E00382: test ax, ax
  loc_00E00385: jz 00E003C4h
  loc_00E00387: mov edx, [esi]
  loc_00E00389: lea eax, var_38
  loc_00E0038C: push eax
  loc_00E0038D: push 006DF0BCh ; "uxtheme.dll"
  loc_00E00392: push esi
  loc_00E00393: call [edx+000009E0h]
  loc_00E00399: mov cx, var_38
  loc_00E0039D: lea edx, var_34
  loc_00E003A0: mov [esi+000000D2h], cx
  loc_00E003A7: mov ecx, [esi+000001E8h]
  loc_00E003AD: add ecx, 00000010h
  loc_00E003B0: mov var_2C, FFFFFFFFh
  loc_00E003B7: mov var_34, 0000000Bh
  loc_00E003BE: call [0040101Ch] ; __vbaVarMove
  loc_00E003C4: mov dx, [esi+000000D2h]
  loc_00E003CB: push 00E003E0h
  loc_00E003D0: mov var_14, edx
  loc_00E003D3: jmp 00E003DFh
  loc_00E003D5: lea ecx, var_24
  loc_00E003D8: call [00401024h] ; __vbaFreeVar
  loc_00E003DE: ret
  loc_00E003DF: ret
  loc_00E003E0: mov eax, arg_C
  loc_00E003E3: mov cx, var_14
  loc_00E003E7: pop edi
  loc_00E003E8: pop esi
  loc_00E003E9: mov [eax], cx
  loc_00E003EC: mov ecx, var_10
  loc_00E003EF: xor eax, eax
  loc_00E003F1: mov fs:[00000000h], ecx
  loc_00E003F8: pop ebx
  loc_00E003F9: mov esp, ebp
  loc_00E003FB: pop ebp
  loc_00E003FC: retn 0008h
End Sub

Private Function Proc_2_146_E00400(arg_C, arg_10, arg_14) 'E00400
  loc_00E00400: push ebp
  loc_00E00401: mov ebp, esp
  loc_00E00403: sub esp, 00000008h
  loc_00E00406: push 00402836h ; __vbaExceptHandler
  loc_00E0040B: mov eax, fs:[00000000h]
  loc_00E00411: push eax
  loc_00E00412: mov fs:[00000000h], esp
  loc_00E00419: sub esp, 0000001Ch
  loc_00E0041C: push ebx
  loc_00E0041D: push esi
  loc_00E0041E: push edi
  loc_00E0041F: mov var_8, esp
  loc_00E00422: mov var_4, 00401B90h
  loc_00E00429: mov edx, arg_C
  loc_00E0042C: mov esi, [004011B0h] ; __vbaStrCopy
  loc_00E00432: xor eax, eax
  loc_00E00434: lea ecx, var_1C
  loc_00E00437: mov var_18, eax
  loc_00E0043A: mov var_1C, eax
  loc_00E0043D: mov var_20, eax
  loc_00E00440: mov var_24, eax
  loc_00E00443: call __vbaStrCopy
  loc_00E00445: mov edx, arg_10
  loc_00E00448: lea ecx, var_18
  loc_00E0044B: call __vbaStrCopy
  loc_00E0044D: mov eax, var_18
  loc_00E00450: lea ecx, var_24
  loc_00E00453: push eax
  loc_00E00454: push ecx
  loc_00E00455: call [004011ECh] ; __vbaStrToAnsi
  loc_00E0045B: push eax
  loc_00E0045C: call 006DE444h ; GetModuleHandle(%x1v)
  loc_00E00461: mov edi, [00401074h] ; __vbaSetSystemError
  loc_00E00467: mov esi, eax
  loc_00E00469: call edi
  loc_00E0046B: mov edx, var_24
  loc_00E0046E: lea eax, var_18
  loc_00E00471: push edx
  loc_00E00472: push eax
  loc_00E00473: call [00401160h] ; __vbaStrToUnicode
  loc_00E00479: lea ecx, var_24
  loc_00E0047C: call [00401258h] ; __vbaFreeStr
  loc_00E00482: test esi, esi
  loc_00E00484: jnz 00E004B9h
  loc_00E00486: mov ecx, var_18
  loc_00E00489: lea edx, var_24
  loc_00E0048C: push ecx
  loc_00E0048D: push edx
  loc_00E0048E: call [004011ECh] ; __vbaStrToAnsi
  loc_00E00494: push eax
  loc_00E00495: call 006DE4D0h ; LoadLibrary(%x1v)
  loc_00E0049A: mov esi, eax
  loc_00E0049C: call edi
  loc_00E0049E: mov eax, var_24
  loc_00E004A1: lea ecx, var_18
  loc_00E004A4: push eax
  loc_00E004A5: push ecx
  loc_00E004A6: call [00401160h] ; __vbaStrToUnicode
  loc_00E004AC: lea ecx, var_24
  loc_00E004AF: call [00401258h] ; __vbaFreeStr
  loc_00E004B5: test esi, esi
  loc_00E004B7: jz 00E004FDh
  loc_00E004B9: mov edx, var_1C
  loc_00E004BC: lea eax, var_24
  loc_00E004BF: push edx
  loc_00E004C0: push eax
  loc_00E004C1: call [004011ECh] ; __vbaStrToAnsi
  loc_00E004C7: push eax
  loc_00E004C8: push esi
  loc_00E004C9: call 006DE5C0h ; GetProcAddress(%x1v, %x2v)
  loc_00E004CE: mov ebx, eax
  loc_00E004D0: call edi
  loc_00E004D2: mov ecx, var_24
  loc_00E004D5: lea edx, var_1C
  loc_00E004D8: push ecx
  loc_00E004D9: push edx
  loc_00E004DA: call [00401160h] ; __vbaStrToUnicode
  loc_00E004E0: xor eax, eax
  loc_00E004E2: lea ecx, var_24
  loc_00E004E5: test ebx, ebx
  loc_00E004E7: setnz al
  loc_00E004EA: neg eax
  loc_00E004EC: mov var_20, eax
  loc_00E004EF: call [00401258h] ; __vbaFreeStr
  loc_00E004F5: push esi
  loc_00E004F6: call 006DE488h ; FreeLibrary(%x1v)
  loc_00E004FB: call edi
  loc_00E004FD: push 00E0051Fh
  loc_00E00502: jmp 00E0050Eh
  loc_00E00504: lea ecx, var_24
  loc_00E00507: call [00401258h] ; __vbaFreeStr
  loc_00E0050D: ret
  loc_00E0050E: mov esi, [00401258h] ; __vbaFreeStr
  loc_00E00514: lea ecx, var_18
  loc_00E00517: call __vbaFreeStr
  loc_00E00519: lea ecx, var_1C
  loc_00E0051C: call __vbaFreeStr
  loc_00E0051E: ret
  loc_00E0051F: mov ecx, arg_14
  loc_00E00522: mov dx, var_20
  loc_00E00526: pop edi
  loc_00E00527: pop esi
  loc_00E00528: mov [ecx], dx
  loc_00E0052B: mov ecx, var_10
  loc_00E0052E: xor eax, eax
  loc_00E00530: mov fs:[00000000h], ecx
  loc_00E00537: pop ebx
  loc_00E00538: mov esp, ebp
  loc_00E0053A: pop ebp
  loc_00E0053B: retn 0010h
End Function

Private Sub Proc_2_147_E00540(arg_C, arg_10) 'E00540
  loc_00E00540: sub esp, 00000010h
  loc_00E00543: mov ecx, arg_C
  loc_00E00547: xor eax, eax
  loc_00E00549: mov [esp], eax
  loc_00E0054D: mov var_4, eax
  loc_00E00551: mov [esp+00000008h], eax
  loc_00E00555: cmp [ecx+00000040h], ax
  loc_00E00559: mov Proc_2_147_E00540, eax
  loc_00E0055D: jz 00E00587h
  loc_00E0055F: mov edx, arg_10
  loc_00E00563: lea eax, [esp]
  loc_00E00567: push eax
  loc_00E00568: mov var_4, 00000010h
  loc_00E00570: mov [esp+00000008h], 00000002h
  loc_00E00578: mov Proc_2_147_E00540, edx
  loc_00E0057C: call 006DE518h ; TrackMouseEvent()
  loc_00E00581: call [00401074h] ; __vbaSetSystemError
  loc_00E00587: xor eax, eax
  loc_00E00589: add esp, 00000010h
  loc_00E0058C: retn 0008h
End Sub

Public Sub mFont_FontChanged(PropertyName) 'E01760
  loc_00E01760: push ebp
  loc_00E01761: mov ebp, esp
  loc_00E01763: sub esp, 0000000Ch
  loc_00E01766: push 00402836h ; __vbaExceptHandler
  loc_00E0176B: mov eax, fs:[00000000h]
  loc_00E01771: push eax
  loc_00E01772: mov fs:[00000000h], esp
  loc_00E01779: sub esp, 00000030h
  loc_00E0177C: push ebx
  loc_00E0177D: push esi
  loc_00E0177E: push edi
  loc_00E0177F: mov var_C, esp
  loc_00E01782: mov var_8, 00401C98h
  loc_00E01789: mov esi, Me
  loc_00E0178C: mov eax, esi
  loc_00E0178E: and eax, 00000001h
  loc_00E01791: mov var_4, eax
  loc_00E01794: and esi, FFFFFFFEh
  loc_00E01797: push esi
  loc_00E01798: mov Me, esi
  loc_00E0179B: mov ecx, [esi]
  loc_00E0179D: call [ecx+00000004h]
  loc_00E017A0: mov edx, PropertyName
  loc_00E017A3: xor ebx, ebx
  loc_00E017A5: lea ecx, var_18
  loc_00E017A8: mov var_18, ebx
  loc_00E017AB: mov var_1C, ebx
  loc_00E017AE: mov var_20, ebx
  loc_00E017B1: call [004011B0h] ; __vbaStrCopy
  loc_00E017B7: mov edx, [esi]
  loc_00E017B9: lea eax, var_1C
  loc_00E017BC: push eax
  loc_00E017BD: push esi
  loc_00E017BE: call [edx+00000A2Ch]
  loc_00E017C4: cmp eax, ebx
  loc_00E017C6: fnclex
  loc_00E017C8: jge 00E017E0h
  loc_00E017CA: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E017D0: push 00000A2Ch
  loc_00E017D5: push 006DD090h
  loc_00E017DA: push esi
  loc_00E017DB: push eax
  loc_00E017DC: call edi
  loc_00E017DE: jmp 00E017E6h
  loc_00E017E0: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E017E6: mov eax, var_1C
  loc_00E017E9: mov ecx, [esi+00000010h]
  loc_00E017EC: lea edx, var_20
  loc_00E017EF: push eax
  loc_00E017F0: mov var_1C, ebx
  loc_00E017F3: mov ebx, [ecx]
  loc_00E017F5: push edx
  loc_00E017F6: call [004010ACh] ; __vbaObjSet
  loc_00E017FC: push eax
  loc_00E017FD: mov eax, [esi+00000010h]
  loc_00E01800: push eax
  loc_00E01801: call [ebx+0000024Ch]
  loc_00E01807: test eax, eax
  loc_00E01809: fnclex
  loc_00E0180B: jge 00E0181Eh
  loc_00E0180D: mov ecx, [esi+00000010h]
  loc_00E01810: push 0000024Ch
  loc_00E01815: push 006DCB00h
  loc_00E0181A: push ecx
  loc_00E0181B: push eax
  loc_00E0181C: call edi
  loc_00E0181E: lea ecx, var_20
  loc_00E01821: call [00401254h] ; __vbaFreeObj
  loc_00E01827: mov edx, [esi]
  loc_00E01829: push esi
  loc_00E0182A: call [edx+00000338h]
  loc_00E01830: test eax, eax
  loc_00E01832: fnclex
  loc_00E01834: jge 00E01844h
  loc_00E01836: push 00000338h
  loc_00E0183B: push 006DCB00h
  loc_00E01840: push esi
  loc_00E01841: push eax
  loc_00E01842: call edi
  loc_00E01844: mov eax, [esi]
  loc_00E01846: push esi
  loc_00E01847: call [eax+00000910h]
  loc_00E0184D: sub esp, 00000010h
  loc_00E01850: mov ecx, 00000008h
  loc_00E01855: mov ebx, esp
  loc_00E01857: mov edx, [esi]
  loc_00E01859: mov eax, 006DEBACh ; "Font"
  loc_00E0185E: push esi
  loc_00E0185F: mov [ebx], ecx
  loc_00E01861: mov ecx, var_2C
  loc_00E01864: mov [ebx+00000004h], ecx
  loc_00E01867: mov [ebx+00000008h], eax
  loc_00E0186A: mov eax, var_24
  loc_00E0186D: mov [ebx+0000000Ch], eax
  loc_00E01870: call [edx+00000390h]
  loc_00E01876: test eax, eax
  loc_00E01878: fnclex
  loc_00E0187A: jge 00E0188Ah
  loc_00E0187C: push 00000390h
  loc_00E01881: push 006DCB00h
  loc_00E01886: push esi
  loc_00E01887: push eax
  loc_00E01888: call edi
  loc_00E0188A: mov var_4, 00000000h
  loc_00E01891: push 00E018B6h
  loc_00E01896: jmp 00E018ACh
  loc_00E01898: lea ecx, var_20
  loc_00E0189B: lea edx, var_1C
  loc_00E0189E: push ecx
  loc_00E0189F: push edx
  loc_00E018A0: push 00000002h
  loc_00E018A2: call [00401048h] ; __vbaFreeObjList
  loc_00E018A8: add esp, 0000000Ch
  loc_00E018AB: ret
  loc_00E018AC: lea ecx, var_18
  loc_00E018AF: call [00401258h] ; __vbaFreeStr
  loc_00E018B5: ret
  loc_00E018B6: mov eax, Me
  loc_00E018B9: push eax
  loc_00E018BA: mov ecx, [eax]
  loc_00E018BC: call [ecx+00000008h]
  loc_00E018BF: mov eax, var_4
  loc_00E018C2: mov ecx, var_14
  loc_00E018C5: pop edi
  loc_00E018C6: pop esi
  loc_00E018C7: mov fs:[00000000h], ecx
  loc_00E018CE: pop ebx
  loc_00E018CF: mov esp, ebp
  loc_00E018D1: pop ebp
  loc_00E018D2: retn 0008h
End Sub

Private Function Proc_2_149_E04000(arg_C, arg_10, arg_14) 'E04000
  loc_00E04000: push ebp
  loc_00E04001: mov ebp, esp
  loc_00E04003: sub esp, 00000008h
  loc_00E04006: push 00402836h ; __vbaExceptHandler
  loc_00E0400B: mov eax, fs:[00000000h]
  loc_00E04011: push eax
  loc_00E04012: mov fs:[00000000h], esp
  loc_00E04019: sub esp, 00000020h
  loc_00E0401C: push ebx
  loc_00E0401D: push esi
  loc_00E0401E: push edi
  loc_00E0401F: mov var_8, esp
  loc_00E04022: mov var_4, 00401F10h
  loc_00E04029: mov edx, arg_C
  loc_00E0402C: mov esi, [004011B0h] ; __vbaStrCopy
  loc_00E04032: xor eax, eax
  loc_00E04034: lea ecx, var_18
  loc_00E04037: mov var_14, eax
  loc_00E0403A: mov var_18, eax
  loc_00E0403D: mov var_20, eax
  loc_00E04040: mov var_24, eax
  loc_00E04043: call __vbaStrCopy
  loc_00E04045: mov edx, arg_10
  loc_00E04048: lea ecx, var_14
  loc_00E0404B: call __vbaStrCopy
  loc_00E0404D: mov eax, var_18
  loc_00E04050: mov esi, [004011ECh] ; __vbaStrToAnsi
  loc_00E04056: lea ecx, var_20
  loc_00E04059: push eax
  loc_00E0405A: push ecx
  loc_00E0405B: call __vbaStrToAnsi
  loc_00E0405D: push eax
  loc_00E0405E: call 006DE444h ; GetModuleHandle(%x1v)
  loc_00E04063: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00E04069: mov edi, eax
  loc_00E0406B: call ebx
  loc_00E0406D: mov edx, var_20
  loc_00E04070: lea eax, var_18
  loc_00E04073: push edx
  loc_00E04074: push eax
  loc_00E04075: call [00401160h] ; __vbaStrToUnicode
  loc_00E0407B: mov ecx, var_14
  loc_00E0407E: lea edx, var_24
  loc_00E04081: push ecx
  loc_00E04082: push edx
  loc_00E04083: call __vbaStrToAnsi
  loc_00E04085: push eax
  loc_00E04086: push edi
  loc_00E04087: call 006DE5C0h ; GetProcAddress(%x1v, %x2v)
  loc_00E0408C: mov esi, eax
  loc_00E0408E: call ebx
  loc_00E04090: mov eax, var_24
  loc_00E04093: lea ecx, var_14
  loc_00E04096: push eax
  loc_00E04097: push ecx
  loc_00E04098: call [00401160h] ; __vbaStrToUnicode
  loc_00E0409E: lea edx, var_24
  loc_00E040A1: lea eax, var_20
  loc_00E040A4: push edx
  loc_00E040A5: push eax
  loc_00E040A6: push 00000002h
  loc_00E040A8: mov var_1C, esi
  loc_00E040AB: call [004011BCh] ; __vbaFreeStrList
  loc_00E040B1: add esp, 0000000Ch
  loc_00E040B4: push 00E040E0h
  loc_00E040B9: jmp 00E040CFh
  loc_00E040BB: lea ecx, var_24
  loc_00E040BE: lea edx, var_20
  loc_00E040C1: push ecx
  loc_00E040C2: push edx
  loc_00E040C3: push 00000002h
  loc_00E040C5: call [004011BCh] ; __vbaFreeStrList
  loc_00E040CB: add esp, 0000000Ch
  loc_00E040CE: ret
  loc_00E040CF: mov esi, [00401258h] ; __vbaFreeStr
  loc_00E040D5: lea ecx, var_14
  loc_00E040D8: call __vbaFreeStr
  loc_00E040DA: lea ecx, var_18
  loc_00E040DD: call __vbaFreeStr
  loc_00E040DF: ret
  loc_00E040E0: mov eax, arg_14
  loc_00E040E3: mov ecx, var_1C
  loc_00E040E6: pop edi
  loc_00E040E7: pop esi
  loc_00E040E8: mov [eax], ecx
  loc_00E040EA: mov ecx, var_10
  loc_00E040ED: xor eax, eax
  loc_00E040EF: mov fs:[00000000h], ecx
  loc_00E040F6: pop ebx
  loc_00E040F7: mov esp, ebp
  loc_00E040F9: pop ebp
  loc_00E040FA: retn 0010h
End Function

Private Sub Proc_2_150_E04100(arg_10, arg_14, arg_18) 'E04100
  loc_00E04100: push ebx
  loc_00E04101: mov ebx, [esp+00000008h]
  loc_00E04105: push ebp
  loc_00E04106: push esi
  loc_00E04107: mov eax, [ebx+0000003Ch]
  loc_00E0410A: push edi
  loc_00E0410B: push eax
  loc_00E0410C: push 00000001h
  loc_00E0410E: call [0040117Ch] ; __vbaUbound
  loc_00E04114: mov edi, eax
  loc_00E04116: test edi, edi
  loc_00E04118: jl 00E041D3h
  loc_00E0411E: mov ebp, [00401100h] ; __vbaGenerateBoundsError
  loc_00E04124: mov eax, [ebx+0000003Ch]
  loc_00E04127: test eax, eax
  loc_00E04129: jz 00E0414Fh
  loc_00E0412B: cmp [eax], 0001h
  loc_00E0412F: jnz 00E0414Fh
  loc_00E04131: mov edx, [eax+00000014h]
  loc_00E04134: mov ecx, [eax+00000010h]
  loc_00E04137: mov esi, edi
  loc_00E04139: sub esi, edx
  loc_00E0413B: cmp esi, ecx
  loc_00E0413D: jb 00E04141h
  loc_00E0413F: call ebp
  loc_00E04141: lea eax, [esi*8]
  loc_00E04148: sub eax, esi
  loc_00E0414A: shl eax, 02h
  loc_00E0414D: jmp 00E04151h
  loc_00E0414F: call ebp
  loc_00E04151: mov ecx, [ebx+0000003Ch]
  loc_00E04154: mov esi, arg_10
  loc_00E04158: mov edx, [ecx+0000000Ch]
  loc_00E0415B: cmp [edx+eax], esi
  loc_00E0415E: jnz 00E0416Ah
  loc_00E04160: cmp arg_14, 0000h
  loc_00E04166: jz 00E041C4h
  loc_00E04168: jmp 00E041A8h
  loc_00E0416A: test ecx, ecx
  loc_00E0416C: jz 00E04192h
  loc_00E0416E: cmp [ecx], 0001h
  loc_00E04172: jnz 00E04192h
  loc_00E04174: mov edx, [ecx+00000014h]
  loc_00E04177: mov eax, [ecx+00000010h]
  loc_00E0417A: mov esi, edi
  loc_00E0417C: sub esi, edx
  loc_00E0417E: cmp esi, eax
  loc_00E04180: jb 00E04184h
  loc_00E04182: call ebp
  loc_00E04184: lea eax, [esi*8]
  loc_00E0418B: sub eax, esi
  loc_00E0418D: shl eax, 02h
  loc_00E04190: jmp 00E04194h
  loc_00E04192: call ebp
  loc_00E04194: mov ecx, [ebx+0000003Ch]
  loc_00E04197: mov edx, [ecx+0000000Ch]
  loc_00E0419A: cmp [edx+eax], 00000000h
  loc_00E0419E: jnz 00E041A8h
  loc_00E041A0: cmp arg_14, 0000h
  loc_00E041A6: jnz 00E041B5h
  loc_00E041A8: sub edi, 00000001h
  loc_00E041AB: jo 00E041E2h
  loc_00E041AD: test edi, edi
  loc_00E041AF: jge 00E04124h
  loc_00E041B5: mov eax, arg_18
  loc_00E041B9: mov [eax], edi
  loc_00E041BB: pop edi
  loc_00E041BC: pop esi
  loc_00E041BD: pop ebp
  loc_00E041BE: xor eax, eax
  loc_00E041C0: pop ebx
  loc_00E041C1: retn 0010h
End Sub

Private Sub Proc_2_151_E041F0
  loc_00E041F0: mov eax, [esp+00000008h]
  loc_00E041F4: mov [eax], 0000h
  loc_00E041F9: xor eax, eax
  loc_00E041FB: retn 0008h
End Sub

Private Sub Proc_2_152_E04200(arg_C, arg_10) 'E04200
  loc_00E04200: push ebp
  loc_00E04201: mov ebp, esp
  loc_00E04203: sub esp, 00000008h
  loc_00E04206: push 00402836h ; __vbaExceptHandler
  loc_00E0420B: mov eax, fs:[00000000h]
  loc_00E04211: push eax
  loc_00E04212: mov fs:[00000000h], esp
  loc_00E04219: sub esp, 00000078h
  loc_00E0421C: push ebx
  loc_00E0421D: push esi
  loc_00E0421E: push edi
  loc_00E0421F: mov var_8, esp
  loc_00E04222: mov var_4, 00401F20h
  loc_00E04229: mov esi, Me
  loc_00E0422C: xor ebx, ebx
  loc_00E0422E: mov var_18, ebx
  loc_00E04231: mov var_24, ebx
  loc_00E04234: mov eax, [esi+000001E8h]
  loc_00E0423A: mov var_28, ebx
  loc_00E0423D: mov var_38, ebx
  loc_00E04240: mov var_48, ebx
  loc_00E04243: mov var_58, ebx
  loc_00E04246: mov var_68, ebx
  loc_00E04249: mov var_6C, ebx
  loc_00E0424C: mov var_70, ebx
  loc_00E0424F: mov var_78, ebx
  loc_00E04252: mov ecx, [eax+00000034h]
  loc_00E04255: mov var_14, ebx
  loc_00E04258: mov var_20, ebx
  loc_00E0425B: cmp [ecx], 00h
  loc_00E0425E: jz 00E042BAh
  loc_00E04260: mov ecx, arg_C
  loc_00E04263: mov edx, [esi]
  loc_00E04265: lea eax, var_70
  loc_00E04268: push eax
  loc_00E04269: push FFFFFFFFh
  loc_00E0426B: push ecx
  loc_00E0426C: push esi
  loc_00E0426D: call [edx+000009FCh]
  loc_00E04273: mov eax, var_70
  loc_00E04276: cmp eax, FFFFFFFFh
  loc_00E04279: mov var_20, eax
  loc_00E0427C: jnz 00E042AFh
  loc_00E0427E: mov edx, [esi+0000003Ch]
  loc_00E04281: lea edi, [esi+0000003Ch]
  loc_00E04284: push edx
  loc_00E04285: push 00000001h
  loc_00E04287: call [0040117Ch] ; __vbaUbound
  loc_00E0428D: add eax, 00000001h
  loc_00E04290: push ebx
  loc_00E04291: jo 00E0466Dh
  loc_00E04297: push eax
  loc_00E04298: push 00000001h
  loc_00E0429A: push 006DD168h ; UDT_2_006DD168
  loc_00E0429F: push edi
  loc_00E042A0: push 0000001Ch
  loc_00E042A2: push ebx
  loc_00E042A3: mov var_20, eax
  loc_00E042A6: call [00401124h] ; __vbaRedimPreserve
  loc_00E042AC: add esp, 0000001Ch
  loc_00E042AF: mov eax, var_20
  loc_00E042B2: mov var_14, eax
  loc_00E042B5: jmp 00E044C2h
  loc_00E042BA: mov edx, 006DF338h ; "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
  loc_00E042BF: lea ecx, var_18
  loc_00E042C2: call [004011B0h] ; __vbaStrCopy
  loc_00E042C8: mov edi, 00000001h
  loc_00E042CD: mov eax, 000000C8h
  loc_00E042D2: cmp edi, eax
  loc_00E042D4: jg 00E043D3h
  loc_00E042DA: lea edx, var_58
  loc_00E042DD: push 00000002h
  loc_00E042DF: lea eax, var_38
  loc_00E042E2: lea ecx, var_18
  loc_00E042E5: push edx
  loc_00E042E6: push eax
  loc_00E042E7: mov var_60, 006DF2E8h ; "&H"
  loc_00E042EE: mov var_68, 00000008h
  loc_00E042F5: mov var_50, ecx
  loc_00E042F8: mov var_58, 00004008h
  loc_00E042FF: call [0040121Ch] ; rtcLeftCharVar
  loc_00E04305: lea ebx, [edi-00000001h]
  loc_00E04308: cmp ebx, 000000C8h
  loc_00E0430E: jb 00E04316h
  loc_00E04310: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04316: lea ecx, var_68
  loc_00E04319: lea edx, var_38
  loc_00E0431C: push ecx
  loc_00E0431D: lea eax, var_48
  loc_00E04320: push edx
  loc_00E04321: push eax
  loc_00E04322: call [00401184h] ; __vbaVarCat
  loc_00E04328: lea ecx, var_24
  loc_00E0432B: push eax
  loc_00E0432C: push ecx
  loc_00E0432D: call [00401180h] ; __vbaStrVarVal
  loc_00E04333: push eax
  loc_00E04334: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E0433A: call [0040111Ch] ; __vbaFpUI1
  loc_00E04340: mov edx, [esi+000001E8h]
  loc_00E04346: mov ecx, [edx+00000034h]
  loc_00E04349: mov [ecx+ebx], al
  loc_00E0434C: lea ecx, var_24
  loc_00E0434F: call [00401258h] ; __vbaFreeStr
  loc_00E04355: mov ebx, [00401038h] ; __vbaFreeVarList
  loc_00E0435B: lea edx, var_48
  loc_00E0435E: lea eax, var_38
  loc_00E04361: push edx
  loc_00E04362: push eax
  loc_00E04363: push 00000002h
  loc_00E04365: call ebx
  loc_00E04367: add esp, 0000000Ch
  loc_00E0436A: lea ecx, var_18
  loc_00E0436D: lea edx, var_38
  loc_00E04370: mov var_50, ecx
  loc_00E04373: push edx
  loc_00E04374: lea eax, var_58
  loc_00E04377: push 00000003h
  loc_00E04379: lea ecx, var_48
  loc_00E0437C: push eax
  loc_00E0437D: push ecx
  loc_00E0437E: mov var_30, 80020004h
  loc_00E04385: mov var_38, 0000000Ah
  loc_00E0438C: mov var_58, 00004008h
  loc_00E04393: call [004010F0h] ; rtcMidCharVar
  loc_00E04399: lea edx, var_48
  loc_00E0439C: push edx
  loc_00E0439D: call [00401034h] ; __vbaStrVarMove
  loc_00E043A3: mov edx, eax
  loc_00E043A5: lea ecx, var_18
  loc_00E043A8: call [00401228h] ; __vbaStrMove
  loc_00E043AE: lea eax, var_48
  loc_00E043B1: lea ecx, var_38
  loc_00E043B4: push eax
  loc_00E043B5: push ecx
  loc_00E043B6: push 00000002h
  loc_00E043B8: call ebx
  loc_00E043BA: mov eax, 00000001h
  loc_00E043BF: add esp, 0000000Ch
  loc_00E043C2: add eax, edi
  loc_00E043C4: jo 00E0466Dh
  loc_00E043CA: mov edi, eax
  loc_00E043CC: xor ebx, ebx
  loc_00E043CE: jmp 00E042CDh
  loc_00E043D3: mov edx, [esi]
  loc_00E043D5: lea eax, var_6C
  loc_00E043D8: push eax
  loc_00E043D9: push esi
  loc_00E043DA: call [edx+00000A00h]
  loc_00E043E0: cmp var_6C, bx
  loc_00E043E4: jz 00E04463h
  loc_00E043E6: mov edi, [00401144h] ; __vbaUI1I2
  loc_00E043EC: mov ecx, 00000090h
  loc_00E043F1: call edi
  loc_00E043F3: mov ecx, [esi+000001E8h]
  loc_00E043F9: mov edx, [ecx+00000034h]
  loc_00E043FC: mov ecx, 00000090h
  loc_00E04401: mov [edx+0000000Fh], al
  loc_00E04404: call edi
  loc_00E04406: mov ecx, [esi+000001E8h]
  loc_00E0440C: mov edx, [ecx+00000034h]
  loc_00E0440F: lea ecx, var_70
  loc_00E04412: push ecx
  loc_00E04413: push 006DF290h ; "EbMode"
  loc_00E04418: mov [edx+00000010h], al
  loc_00E0441B: mov eax, [esi]
  loc_00E0441D: push 006DF2D8h ; "vba6"
  loc_00E04422: push esi
  loc_00E04423: call [eax+000009F8h]
  loc_00E04429: mov edx, [esi+000001E8h]
  loc_00E0442F: mov eax, var_70
  loc_00E04432: mov [edx+00000044h], eax
  loc_00E04435: mov ecx, [esi+000001E8h]
  loc_00E0443B: cmp [ecx+00000044h], ebx
  loc_00E0443E: jnz 00E04463h
  loc_00E04440: mov edx, [esi]
  loc_00E04442: lea eax, var_70
  loc_00E04445: push eax
  loc_00E04446: push 006DF290h ; "EbMode"
  loc_00E0444B: push 006DF2C8h ; "vba5"
  loc_00E04450: push esi
  loc_00E04451: call [edx+000009F8h]
  loc_00E04457: mov ecx, [esi+000001E8h]
  loc_00E0445D: mov edx, var_70
  loc_00E04460: mov [ecx+00000044h], edx
  loc_00E04463: mov eax, [esi]
  loc_00E04465: lea ecx, var_70
  loc_00E04468: push ecx
  loc_00E04469: push 006DF26Ch ; "CallWindowProcA"
  loc_00E0446E: push 006DF020h ; "User32"
  loc_00E04473: push esi
  loc_00E04474: call [eax+000009F8h]
  loc_00E0447A: mov edx, [esi+000001E8h]
  loc_00E04480: mov eax, var_70
  loc_00E04483: mov [edx+00000040h], eax
  loc_00E04486: mov ecx, [esi]
  loc_00E04488: lea edx, var_70
  loc_00E0448B: push edx
  loc_00E0448C: push 006DF2A4h ; "SetWindowLongA"
  loc_00E04491: push 006DF020h ; "User32"
  loc_00E04496: push esi
  loc_00E04497: call [ecx+000009F8h]
  loc_00E0449D: mov eax, [esi+000001E8h]
  loc_00E044A3: mov ecx, var_70
  loc_00E044A6: push ebx
  loc_00E044A7: push ebx
  loc_00E044A8: push 00000001h
  loc_00E044AA: lea edx, [esi+0000003Ch]
  loc_00E044AD: push 006DD168h ; UDT_2_006DD168
  loc_00E044B2: push edx
  loc_00E044B3: push 0000001Ch
  loc_00E044B5: push ebx
  loc_00E044B6: mov [eax+00000048h], ecx
  loc_00E044B9: call [00401138h] ; __vbaRedim
  loc_00E044BF: add esp, 0000001Ch
  loc_00E044C2: mov eax, [esi+0000003Ch]
  loc_00E044C5: lea ecx, var_78
  loc_00E044C8: push eax
  loc_00E044C9: push ecx
  loc_00E044CA: call [004011DCh] ; __vbaAryLock
  loc_00E044D0: mov ecx, var_78
  loc_00E044D3: cmp ecx, ebx
  loc_00E044D5: jz 00E04503h
  loc_00E044D7: cmp [ecx], 0001h
  loc_00E044DB: jnz 00E04503h
  loc_00E044DD: mov edi, var_20
  loc_00E044E0: mov edx, [ecx+00000014h]
  loc_00E044E3: mov eax, [ecx+00000010h]
  loc_00E044E6: sub edi, edx
  loc_00E044E8: cmp edi, eax
  loc_00E044EA: jb 00E044F5h
  loc_00E044EC: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E044F2: mov ecx, var_78
  loc_00E044F5: lea eax, [edi*8]
  loc_00E044FC: sub eax, edi
  loc_00E044FE: shl eax, 02h
  loc_00E04501: jmp 00E0450Ch
  loc_00E04503: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04509: mov ecx, var_78
  loc_00E0450C: mov edi, [ecx+0000000Ch]
  loc_00E0450F: mov edx, arg_C
  loc_00E04512: add edi, eax
  loc_00E04514: push 000000C8h
  loc_00E04519: push ebx
  loc_00E0451A: mov [edi], edx
  loc_00E0451C: call 006DE604h ; GlobalAlloc(%x1v, %x2v)
  loc_00E04521: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00E04527: mov var_70, eax
  loc_00E0452A: call ebx
  loc_00E0452C: mov eax, var_70
  loc_00E0452F: mov edx, [edi]
  loc_00E04531: mov [edi+00000004h], eax
  loc_00E04534: mov ecx, var_70
  loc_00E04537: push ecx
  loc_00E04538: push FFFFFFFCh
  loc_00E0453A: push edx
  loc_00E0453B: call 006DE690h ; SetWindowLong(%x1v, %x2v, %x3v)
  loc_00E04540: mov var_70, eax
  loc_00E04543: call ebx
  loc_00E04545: mov eax, var_70
  loc_00E04548: push 000000C8h
  loc_00E0454D: mov [edi+00000008h], eax
  loc_00E04550: mov ecx, [esi+000001E8h]
  loc_00E04556: mov eax, [edi+00000004h]
  loc_00E04559: mov edx, [ecx+00000034h]
  loc_00E0455C: push edx
  loc_00E0455D: push eax
  loc_00E0455E: call 006DE3F8h ; CopyMemory(%x1v, %x2v, %x3v)
  loc_00E04563: call ebx
  loc_00E04565: mov edx, [esi+000001E8h]
  loc_00E0456B: mov ecx, [esi]
  loc_00E0456D: mov eax, [edx+00000044h]
  loc_00E04570: mov edx, [edi+00000004h]
  loc_00E04573: push eax
  loc_00E04574: push 00000012h
  loc_00E04576: push edx
  loc_00E04577: push esi
  loc_00E04578: call [ecx+00000A1Ch]
  loc_00E0457E: mov ecx, [edi+00000008h]
  loc_00E04581: mov edx, [edi+00000004h]
  loc_00E04584: mov eax, [esi]
  loc_00E04586: push ecx
  loc_00E04587: push 00000044h
  loc_00E04589: push edx
  loc_00E0458A: push esi
  loc_00E0458B: call [eax+00000A20h]
  loc_00E04591: mov ecx, [esi+000001E8h]
  loc_00E04597: mov eax, [esi]
  loc_00E04599: mov edx, [ecx+00000048h]
  loc_00E0459C: mov ecx, [edi+00000004h]
  loc_00E0459F: push edx
  loc_00E045A0: push 0000004Eh
  loc_00E045A2: push ecx
  loc_00E045A3: push esi
  loc_00E045A4: call [eax+00000A1Ch]
  loc_00E045AA: mov eax, [edi+00000008h]
  loc_00E045AD: mov ecx, [edi+00000004h]
  loc_00E045B0: mov edx, [esi]
  loc_00E045B2: push eax
  loc_00E045B3: push 00000074h
  loc_00E045B5: push ecx
  loc_00E045B6: push esi
  loc_00E045B7: call [edx+00000A20h]
  loc_00E045BD: mov eax, [esi+000001E8h]
  loc_00E045C3: mov edx, [esi]
  loc_00E045C5: mov ecx, [eax+00000040h]
  loc_00E045C8: mov eax, [edi+00000004h]
  loc_00E045CB: push ecx
  loc_00E045CC: push 00000079h
  loc_00E045CE: push eax
  loc_00E045CF: push esi
  loc_00E045D0: call [edx+00000A1Ch]
  loc_00E045D6: mov ebx, [esi]
  loc_00E045D8: lea ecx, var_28
  loc_00E045DB: push esi
  loc_00E045DC: push ecx
  loc_00E045DD: call [004010B4h] ; __vbaObjSetAddref
  loc_00E045E3: push eax
  loc_00E045E4: call [00401190h] ; VarPtr
  loc_00E045EA: push eax
  loc_00E045EB: mov edx, [edi+00000004h]
  loc_00E045EE: push 000000BAh
  loc_00E045F3: push edx
  loc_00E045F4: push esi
  loc_00E045F5: call [ebx+00000A20h]
  loc_00E045FB: lea ecx, var_28
  loc_00E045FE: call [00401254h] ; __vbaFreeObj
  loc_00E04604: lea eax, var_78
  loc_00E04607: push eax
  loc_00E04608: call [00401244h] ; __vbaAryUnlock
  loc_00E0460E: fwait
  loc_00E0460F: push 00E04650h
  loc_00E04614: jmp 00E0463Ch
  loc_00E04616: lea ecx, var_24
  loc_00E04619: call [00401258h] ; __vbaFreeStr
  loc_00E0461F: lea ecx, var_28
  loc_00E04622: call [00401254h] ; __vbaFreeObj
  loc_00E04628: lea ecx, var_48
  loc_00E0462B: lea edx, var_38
  loc_00E0462E: push ecx
  loc_00E0462F: push edx
  loc_00E04630: push 00000002h
  loc_00E04632: call [00401038h] ; __vbaFreeVarList
  loc_00E04638: add esp, 0000000Ch
  loc_00E0463B: ret
  loc_00E0463C: lea eax, var_78
  loc_00E0463F: push eax
  loc_00E04640: call [00401244h] ; __vbaAryUnlock
  loc_00E04646: lea ecx, var_18
  loc_00E04649: call [00401258h] ; __vbaFreeStr
  loc_00E0464F: ret
  loc_00E04650: mov ecx, arg_10
  loc_00E04653: mov edx, var_14
  loc_00E04656: pop edi
  loc_00E04657: pop esi
  loc_00E04658: mov [ecx], edx
  loc_00E0465A: mov ecx, var_10
  loc_00E0465D: xor eax, eax
  loc_00E0465F: mov fs:[00000000h], ecx
  loc_00E04666: pop ebx
  loc_00E04667: mov esp, ebp
  loc_00E04669: pop ebp
  loc_00E0466A: retn 000Ch
End Sub

Private Sub Proc_2_153_E04680
  loc_00E04680: mov eax, [esp+00000008h]
  loc_00E04684: mov ecx, Proc_2_153_E04680
  loc_00E04688: mov [eax], FFFFFFh
  loc_00E0468D: mov [ecx], FFFFFFh
  loc_00E04692: xor eax, eax
  loc_00E04694: retn 000Ch
End Sub

Private Function Proc_2_154_E046A0(arg_C, arg_10, arg_14) 'E046A0
  loc_00E046A0: push ebp
  loc_00E046A1: mov ebp, esp
  loc_00E046A3: sub esp, 00000008h
  loc_00E046A6: push 00402836h ; __vbaExceptHandler
  loc_00E046AB: mov eax, fs:[00000000h]
  loc_00E046B1: push eax
  loc_00E046B2: mov fs:[00000000h], esp
  loc_00E046B9: sub esp, 00000014h
  loc_00E046BC: push ebx
  loc_00E046BD: push esi
  loc_00E046BE: push edi
  loc_00E046BF: mov var_8, esp
  loc_00E046C2: mov var_4, 00401F30h
  loc_00E046C9: mov edi, Me
  loc_00E046CC: lea ecx, var_1C
  loc_00E046CF: xor esi, esi
  loc_00E046D1: mov eax, [edi+0000003Ch]
  loc_00E046D4: mov var_14, esi
  loc_00E046D7: push eax
  loc_00E046D8: push ecx
  loc_00E046D9: mov var_1C, esi
  loc_00E046DC: call [004011DCh] ; __vbaAryLock
  loc_00E046E2: mov eax, var_1C
  loc_00E046E5: cmp eax, esi
  loc_00E046E7: jz 00E04727h
  loc_00E046E9: cmp [eax], 0001h
  loc_00E046ED: jnz 00E04727h
  loc_00E046EF: mov ecx, arg_C
  loc_00E046F2: mov edx, [edi]
  loc_00E046F4: lea eax, var_14
  loc_00E046F7: push eax
  loc_00E046F8: push esi
  loc_00E046F9: push ecx
  loc_00E046FA: push edi
  loc_00E046FB: call [edx+000009FCh]
  loc_00E04701: mov eax, var_1C
  loc_00E04704: mov esi, var_14
  loc_00E04707: mov edx, [eax+00000014h]
  loc_00E0470A: mov ecx, [eax+00000010h]
  loc_00E0470D: sub esi, edx
  loc_00E0470F: cmp esi, ecx
  loc_00E04711: jb 00E04719h
  loc_00E04713: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04719: lea eax, [esi*8]
  loc_00E04720: sub eax, esi
  loc_00E04722: shl eax, 02h
  loc_00E04725: jmp 00E0472Dh
  loc_00E04727: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E0472D: mov edx, var_1C
  loc_00E04730: mov bl, arg_14
  loc_00E04733: mov esi, [edx+0000000Ch]
  loc_00E04736: add esi, eax
  loc_00E04738: test bl, 02h
  loc_00E0473B: jz 00E04758h
  loc_00E0473D: mov ecx, [esi+00000004h]
  loc_00E04740: mov eax, [edi]
  loc_00E04742: push ecx
  loc_00E04743: lea edx, [esi+00000010h]
  loc_00E04746: push 00000002h
  loc_00E04748: push edx
  loc_00E04749: mov edx, arg_10
  loc_00E0474C: lea ecx, [esi+00000018h]
  loc_00E0474F: push ecx
  loc_00E04750: push edx
  loc_00E04751: push edi
  loc_00E04752: call [eax+00000A14h]
  loc_00E04758: test bl, 01h
  loc_00E0475B: jz 00E04778h
  loc_00E0475D: mov ecx, [esi+00000004h]
  loc_00E04760: mov eax, [edi]
  loc_00E04762: push ecx
  loc_00E04763: mov ecx, arg_10
  loc_00E04766: lea edx, [esi+0000000Ch]
  loc_00E04769: push 00000001h
  loc_00E0476B: add esi, 00000014h
  loc_00E0476E: push edx
  loc_00E0476F: push esi
  loc_00E04770: push ecx
  loc_00E04771: push edi
  loc_00E04772: call [eax+00000A14h]
  loc_00E04778: lea edx, var_1C
  loc_00E0477B: push edx
  loc_00E0477C: call [00401244h] ; __vbaAryUnlock
  loc_00E04782: push 00E04792h
  loc_00E04787: lea eax, var_1C
  loc_00E0478A: push eax
  loc_00E0478B: call [00401244h] ; __vbaAryUnlock
  loc_00E04791: ret
  loc_00E04792: mov ecx, var_10
  loc_00E04795: pop edi
  loc_00E04796: pop esi
  loc_00E04797: xor eax, eax
  loc_00E04799: mov fs:[00000000h], ecx
  loc_00E047A0: pop ebx
  loc_00E047A1: mov esp, ebp
  loc_00E047A3: pop ebp
  loc_00E047A4: retn 0010h
End Function

Private Function Proc_2_155_E047B0(arg_C, arg_10, arg_14) 'E047B0
  loc_00E047B0: push ebp
  loc_00E047B1: mov ebp, esp
  loc_00E047B3: sub esp, 00000008h
  loc_00E047B6: push 00402836h ; __vbaExceptHandler
  loc_00E047BB: mov eax, fs:[00000000h]
  loc_00E047C1: push eax
  loc_00E047C2: mov fs:[00000000h], esp
  loc_00E047C9: sub esp, 00000014h
  loc_00E047CC: push ebx
  loc_00E047CD: push esi
  loc_00E047CE: push edi
  loc_00E047CF: mov var_8, esp
  loc_00E047D2: mov var_4, 00401F40h
  loc_00E047D9: mov edi, Me
  loc_00E047DC: lea ecx, var_1C
  loc_00E047DF: xor esi, esi
  loc_00E047E1: mov eax, [edi+0000003Ch]
  loc_00E047E4: mov var_14, esi
  loc_00E047E7: push eax
  loc_00E047E8: push ecx
  loc_00E047E9: mov var_1C, esi
  loc_00E047EC: call [004011DCh] ; __vbaAryLock
  loc_00E047F2: mov eax, var_1C
  loc_00E047F5: cmp eax, esi
  loc_00E047F7: jz 00E04837h
  loc_00E047F9: cmp [eax], 0001h
  loc_00E047FD: jnz 00E04837h
  loc_00E047FF: mov ecx, arg_C
  loc_00E04802: mov edx, [edi]
  loc_00E04804: lea eax, var_14
  loc_00E04807: push eax
  loc_00E04808: push esi
  loc_00E04809: push ecx
  loc_00E0480A: push edi
  loc_00E0480B: call [edx+000009FCh]
  loc_00E04811: mov eax, var_1C
  loc_00E04814: mov esi, var_14
  loc_00E04817: mov edx, [eax+00000014h]
  loc_00E0481A: mov ecx, [eax+00000010h]
  loc_00E0481D: sub esi, edx
  loc_00E0481F: cmp esi, ecx
  loc_00E04821: jb 00E04829h
  loc_00E04823: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04829: lea eax, [esi*8]
  loc_00E04830: sub eax, esi
  loc_00E04832: shl eax, 02h
  loc_00E04835: jmp 00E0483Dh
  loc_00E04837: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E0483D: mov edx, var_1C
  loc_00E04840: mov bl, arg_14
  loc_00E04843: mov esi, [edx+0000000Ch]
  loc_00E04846: add esi, eax
  loc_00E04848: test bl, 02h
  loc_00E0484B: jz 00E04868h
  loc_00E0484D: mov ecx, [esi+00000004h]
  loc_00E04850: mov eax, [edi]
  loc_00E04852: push ecx
  loc_00E04853: lea edx, [esi+00000010h]
  loc_00E04856: push 00000002h
  loc_00E04858: push edx
  loc_00E04859: mov edx, arg_10
  loc_00E0485C: lea ecx, [esi+00000018h]
  loc_00E0485F: push ecx
  loc_00E04860: push edx
  loc_00E04861: push edi
  loc_00E04862: call [eax+00000A18h]
  loc_00E04868: test bl, 01h
  loc_00E0486B: jz 00E04888h
  loc_00E0486D: mov ecx, [esi+00000004h]
  loc_00E04870: mov eax, [edi]
  loc_00E04872: push ecx
  loc_00E04873: mov ecx, arg_10
  loc_00E04876: lea edx, [esi+0000000Ch]
  loc_00E04879: push 00000001h
  loc_00E0487B: add esi, 00000014h
  loc_00E0487E: push edx
  loc_00E0487F: push esi
  loc_00E04880: push ecx
  loc_00E04881: push edi
  loc_00E04882: call [eax+00000A18h]
  loc_00E04888: lea edx, var_1C
  loc_00E0488B: push edx
  loc_00E0488C: call [00401244h] ; __vbaAryUnlock
  loc_00E04892: push 00E048A2h
  loc_00E04897: lea eax, var_1C
  loc_00E0489A: push eax
  loc_00E0489B: call [00401244h] ; __vbaAryUnlock
  loc_00E048A1: ret
  loc_00E048A2: mov ecx, var_10
  loc_00E048A5: pop edi
  loc_00E048A6: pop esi
  loc_00E048A7: xor eax, eax
  loc_00E048A9: mov fs:[00000000h], ecx
  loc_00E048B0: pop ebx
  loc_00E048B1: mov esp, ebp
  loc_00E048B3: pop ebp
  loc_00E048B4: retn 0010h
End Function

Private Function Proc_2_156_E048C0(arg_C, arg_10, arg_14, arg_18, arg_1C) 'E048C0
  loc_00E048C0: push ebp
  loc_00E048C1: mov ebp, esp
  loc_00E048C3: sub esp, 00000008h
  loc_00E048C6: push 00402836h ; __vbaExceptHandler
  loc_00E048CB: mov eax, fs:[00000000h]
  loc_00E048D1: push eax
  loc_00E048D2: mov fs:[00000000h], esp
  loc_00E048D9: sub esp, 00000024h
  loc_00E048DC: push ebx
  loc_00E048DD: push esi
  loc_00E048DE: push edi
  loc_00E048DF: mov var_8, esp
  loc_00E048E2: mov var_4, 00401F50h
  loc_00E048E9: xor eax, eax
  loc_00E048EB: push eax
  loc_00E048EC: mov var_18, eax
  loc_00E048EF: mov var_1C, eax
  loc_00E048F2: push 00000001h
  loc_00E048F4: push 00000001h
  loc_00E048F6: lea eax, var_18
  loc_00E048F9: push 00000003h
  loc_00E048FB: push eax
  loc_00E048FC: push 00000004h
  loc_00E048FE: push 00000080h
  loc_00E04903: call [00401138h] ; __vbaRedim
  loc_00E04909: mov eax, arg_C
  loc_00E0490C: add esp, 0000001Ch
  loc_00E0490F: cmp eax, FFFFFFFFh
  loc_00E04912: jnz 00E0492Bh
  loc_00E04914: mov ecx, arg_14
  loc_00E04917: mov ebx, [00401100h] ; __vbaGenerateBoundsError
  loc_00E0491D: mov edi, arg_10
  loc_00E04920: mov [ecx], FFFFFFFFh
  loc_00E04926: jmp 00E04A8Fh
  loc_00E0492B: mov esi, arg_14
  loc_00E0492E: mov edi, 00000001h
  loc_00E04933: mov eax, [esi]
  loc_00E04935: sub eax, 00000001h
  loc_00E04938: jo 00E04CC4h
  loc_00E0493E: mov var_30, eax
  loc_00E04941: cmp edi, eax
  loc_00E04943: jg 00E04A25h
  loc_00E04949: mov esi, arg_10
  loc_00E0494C: mov eax, [esi]
  loc_00E0494E: test eax, eax
  loc_00E04950: jz 00E0497Ah
  loc_00E04952: cmp [eax], 0001h
  loc_00E04956: jnz 00E0497Ah
  loc_00E04958: mov edx, [eax+00000014h]
  loc_00E0495B: mov ecx, [eax+00000010h]
  loc_00E0495E: mov ebx, [00401100h] ; __vbaGenerateBoundsError
  loc_00E04964: mov esi, edi
  loc_00E04966: sub esi, edx
  loc_00E04968: cmp esi, ecx
  loc_00E0496A: jb 00E0496Eh
  loc_00E0496C: call ebx
  loc_00E0496E: lea eax, [esi*4]
  loc_00E04975: mov esi, arg_10
  loc_00E04978: jmp 00E04986h
  loc_00E0497A: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04980: mov ebx, [00401100h] ; __vbaGenerateBoundsError
  loc_00E04986: mov ecx, [esi]
  loc_00E04988: mov edx, [ecx+0000000Ch]
  loc_00E0498B: cmp [edx+eax], 00000000h
  loc_00E0498F: jz 00E049E4h
  loc_00E04991: test ecx, ecx
  loc_00E04993: jz 00E049B7h
  loc_00E04995: cmp [ecx], 0001h
  loc_00E04999: jnz 00E049B7h
  loc_00E0499B: mov edx, [ecx+00000014h]
  loc_00E0499E: mov eax, [ecx+00000010h]
  loc_00E049A1: mov esi, edi
  loc_00E049A3: sub esi, edx
  loc_00E049A5: cmp esi, eax
  loc_00E049A7: jb 00E049ABh
  loc_00E049A9: call ebx
  loc_00E049AB: lea eax, [esi*4]
  loc_00E049B2: mov esi, arg_10
  loc_00E049B5: jmp 00E049B9h
  loc_00E049B7: call ebx
  loc_00E049B9: mov ecx, [esi]
  loc_00E049BB: mov edx, [ecx+0000000Ch]
  loc_00E049BE: mov ecx, arg_C
  loc_00E049C1: cmp [edx+eax], ecx
  loc_00E049C4: jz 00E04C84h
  loc_00E049CA: mov esi, arg_14
  loc_00E049CD: mov eax, 00000001h
  loc_00E049D2: add eax, edi
  loc_00E049D4: jo 00E04CC4h
  loc_00E049DA: mov edi, eax
  loc_00E049DC: mov eax, var_30
  loc_00E049DF: jmp 00E04941h
  loc_00E049E4: test ecx, ecx
  loc_00E049E6: jz 00E04A13h
  loc_00E049E8: cmp [ecx], 0001h
  loc_00E049EC: jnz 00E04A13h
  loc_00E049EE: mov edx, [ecx+00000014h]
  loc_00E049F1: mov eax, [ecx+00000010h]
  loc_00E049F4: sub edi, edx
  loc_00E049F6: cmp edi, eax
  loc_00E049F8: jb 00E049FCh
  loc_00E049FA: call ebx
  loc_00E049FC: mov edx, [esi]
  loc_00E049FE: lea eax, [edi*4]
  loc_00E04A05: mov ecx, [edx+0000000Ch]
  loc_00E04A08: mov edx, arg_C
  loc_00E04A0B: mov [ecx+eax], edx
  loc_00E04A0E: jmp 00E04C84h
  loc_00E04A13: call ebx
  loc_00E04A15: mov edx, [esi]
  loc_00E04A17: mov ecx, [edx+0000000Ch]
  loc_00E04A1A: mov edx, arg_C
  loc_00E04A1D: mov [ecx+eax], edx
  loc_00E04A20: jmp 00E04C84h
  loc_00E04A25: mov eax, [esi]
  loc_00E04A27: mov edi, arg_10
  loc_00E04A2A: add eax, 00000001h
  loc_00E04A2D: push 00000001h
  loc_00E04A2F: jo 00E04CC4h
  loc_00E04A35: push eax
  loc_00E04A36: push 00000001h
  loc_00E04A38: push 00000003h
  loc_00E04A3A: push edi
  loc_00E04A3B: push 00000004h
  loc_00E04A3D: push 00000080h
  loc_00E04A42: mov [esi], eax
  loc_00E04A44: call [00401124h] ; __vbaRedimPreserve
  loc_00E04A4A: mov eax, [edi]
  loc_00E04A4C: add esp, 0000001Ch
  loc_00E04A4F: test eax, eax
  loc_00E04A51: jz 00E04A78h
  loc_00E04A53: cmp [eax], 0001h
  loc_00E04A57: jnz 00E04A78h
  loc_00E04A59: mov esi, [esi]
  loc_00E04A5B: mov edx, [eax+00000014h]
  loc_00E04A5E: mov ecx, [eax+00000010h]
  loc_00E04A61: mov ebx, [00401100h] ; __vbaGenerateBoundsError
  loc_00E04A67: sub esi, edx
  loc_00E04A69: cmp esi, ecx
  loc_00E04A6B: jb 00E04A6Fh
  loc_00E04A6D: call ebx
  loc_00E04A6F: lea eax, [esi*4]
  loc_00E04A76: jmp 00E04A84h
  loc_00E04A78: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04A7E: mov ebx, [00401100h] ; __vbaGenerateBoundsError
  loc_00E04A84: mov ecx, [edi]
  loc_00E04A86: mov edx, [ecx+0000000Ch]
  loc_00E04A89: mov ecx, arg_C
  loc_00E04A8C: mov [edx+eax], ecx
  loc_00E04A8F: mov eax, arg_18
  loc_00E04A92: mov ecx, var_18
  loc_00E04A95: cmp eax, 00000002h
  loc_00E04A98: jnz 00E04B14h
  loc_00E04A9A: test ecx, ecx
  loc_00E04A9C: jz 00E04ABEh
  loc_00E04A9E: cmp [ecx], 0001h
  loc_00E04AA2: jnz 00E04ABEh
  loc_00E04AA4: mov esi, [ecx+00000014h]
  loc_00E04AA7: mov eax, [ecx+00000010h]
  loc_00E04AAA: neg esi
  loc_00E04AAC: cmp esi, eax
  loc_00E04AAE: jb 00E04AB5h
  loc_00E04AB0: call ebx
  loc_00E04AB2: mov ecx, var_18
  loc_00E04AB5: lea eax, [esi*4]
  loc_00E04ABC: jmp 00E04AC3h
  loc_00E04ABE: call ebx
  loc_00E04AC0: mov ecx, var_18
  loc_00E04AC3: mov edx, [ecx+0000000Ch]
  loc_00E04AC6: mov [edx+eax], 00000058h
  loc_00E04ACD: mov ecx, var_18
  loc_00E04AD0: test ecx, ecx
  loc_00E04AD2: jz 00E04B03h
  loc_00E04AD4: cmp [ecx], 0001h
  loc_00E04AD8: jnz 00E04B03h
  loc_00E04ADA: mov edx, [ecx+00000014h]
  loc_00E04ADD: mov eax, [ecx+00000010h]
  loc_00E04AE0: mov esi, 00000001h
  loc_00E04AE5: sub esi, edx
  loc_00E04AE7: cmp esi, eax
  loc_00E04AE9: jb 00E04AF0h
  loc_00E04AEB: call ebx
  loc_00E04AED: mov ecx, var_18
  loc_00E04AF0: mov ecx, [ecx+0000000Ch]
  loc_00E04AF3: lea eax, [esi*4]
  loc_00E04AFA: mov [ecx+eax], 0000005Dh
  loc_00E04B01: jmp 00E04B82h
  loc_00E04B03: call ebx
  loc_00E04B05: mov ecx, var_18
  loc_00E04B08: mov ecx, [ecx+0000000Ch]
  loc_00E04B0B: mov [ecx+eax], 0000005Dh
  loc_00E04B12: jmp 00E04B82h
  loc_00E04B14: test ecx, ecx
  loc_00E04B16: jz 00E04B38h
  loc_00E04B18: cmp [ecx], 0001h
  loc_00E04B1C: jnz 00E04B38h
  loc_00E04B1E: mov esi, [ecx+00000014h]
  loc_00E04B21: mov eax, [ecx+00000010h]
  loc_00E04B24: neg esi
  loc_00E04B26: cmp esi, eax
  loc_00E04B28: jb 00E04B2Fh
  loc_00E04B2A: call ebx
  loc_00E04B2C: mov ecx, var_18
  loc_00E04B2F: lea eax, [esi*4]
  loc_00E04B36: jmp 00E04B3Dh
  loc_00E04B38: call ebx
  loc_00E04B3A: mov ecx, var_18
  loc_00E04B3D: mov edx, [ecx+0000000Ch]
  loc_00E04B40: mov [edx+eax], 00000084h
  loc_00E04B47: mov ecx, var_18
  loc_00E04B4A: test ecx, ecx
  loc_00E04B4C: jz 00E04B73h
  loc_00E04B4E: cmp [ecx], 0001h
  loc_00E04B52: jnz 00E04B73h
  loc_00E04B54: mov edx, [ecx+00000014h]
  loc_00E04B57: mov eax, [ecx+00000010h]
  loc_00E04B5A: mov esi, 00000001h
  loc_00E04B5F: sub esi, edx
  loc_00E04B61: cmp esi, eax
  loc_00E04B63: jb 00E04B6Ah
  loc_00E04B65: call ebx
  loc_00E04B67: mov ecx, var_18
  loc_00E04B6A: lea eax, [esi*4]
  loc_00E04B71: jmp 00E04B78h
  loc_00E04B73: call ebx
  loc_00E04B75: mov ecx, var_18
  loc_00E04B78: mov ecx, [ecx+0000000Ch]
  loc_00E04B7B: mov [ecx+eax], 00000089h
  loc_00E04B82: cmp arg_C, FFFFFFFFh
  loc_00E04B86: jz 00E04C36h
  loc_00E04B8C: mov edx, [edi]
  loc_00E04B8E: lea eax, var_1C
  loc_00E04B91: push edx
  loc_00E04B92: push eax
  loc_00E04B93: call [004011DCh] ; __vbaAryLock
  loc_00E04B99: mov ecx, var_1C
  loc_00E04B9C: test ecx, ecx
  loc_00E04B9E: jz 00E04BC5h
  loc_00E04BA0: cmp [ecx], 0001h
  loc_00E04BA4: jnz 00E04BC5h
  loc_00E04BA6: mov edx, [ecx+00000014h]
  loc_00E04BA9: mov eax, [ecx+00000010h]
  loc_00E04BAC: mov esi, 00000001h
  loc_00E04BB1: sub esi, edx
  loc_00E04BB3: cmp esi, eax
  loc_00E04BB5: jb 00E04BBCh
  loc_00E04BB7: call ebx
  loc_00E04BB9: mov ecx, var_1C
  loc_00E04BBC: lea eax, [esi*4]
  loc_00E04BC3: jmp 00E04BCAh
  loc_00E04BC5: call ebx
  loc_00E04BC7: mov ecx, var_1C
  loc_00E04BCA: mov ecx, [ecx+0000000Ch]
  loc_00E04BCD: add ecx, eax
  loc_00E04BCF: push ecx
  loc_00E04BD0: call [00401190h] ; VarPtr
  loc_00E04BD6: lea edx, var_1C
  loc_00E04BD9: mov ebx, eax
  loc_00E04BDB: push edx
  loc_00E04BDC: call [00401244h] ; __vbaAryUnlock
  loc_00E04BE2: mov ecx, var_18
  loc_00E04BE5: test ecx, ecx
  loc_00E04BE7: jz 00E04C0Dh
  loc_00E04BE9: cmp [ecx], 0001h
  loc_00E04BED: jnz 00E04C0Dh
  loc_00E04BEF: mov esi, [ecx+00000014h]
  loc_00E04BF2: mov eax, [ecx+00000010h]
  loc_00E04BF5: neg esi
  loc_00E04BF7: cmp esi, eax
  loc_00E04BF9: jb 00E04C04h
  loc_00E04BFB: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04C01: mov ecx, var_18
  loc_00E04C04: lea eax, [esi*4]
  loc_00E04C0B: jmp 00E04C16h
  loc_00E04C0D: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04C13: mov ecx, var_18
  loc_00E04C16: mov ecx, [ecx+0000000Ch]
  loc_00E04C19: mov edi, Me
  loc_00E04C1C: push ebx
  loc_00E04C1D: mov eax, [ecx+eax]
  loc_00E04C20: mov ecx, arg_1C
  loc_00E04C23: mov edx, [edi]
  loc_00E04C25: push eax
  loc_00E04C26: push ecx
  loc_00E04C27: push edi
  loc_00E04C28: call [edx+00000A20h]
  loc_00E04C2E: mov ebx, [00401100h] ; __vbaGenerateBoundsError
  loc_00E04C34: jmp 00E04C39h
  loc_00E04C36: mov edi, Me
  loc_00E04C39: mov ecx, var_18
  loc_00E04C3C: test ecx, ecx
  loc_00E04C3E: jz 00E04C65h
  loc_00E04C40: cmp [ecx], 0001h
  loc_00E04C44: jnz 00E04C65h
  loc_00E04C46: mov edx, [ecx+00000014h]
  loc_00E04C49: mov eax, [ecx+00000010h]
  loc_00E04C4C: mov esi, 00000001h
  loc_00E04C51: sub esi, edx
  loc_00E04C53: cmp esi, eax
  loc_00E04C55: jb 00E04C5Ch
  loc_00E04C57: call ebx
  loc_00E04C59: mov ecx, var_18
  loc_00E04C5C: lea eax, [esi*4]
  loc_00E04C63: jmp 00E04C6Ah
  loc_00E04C65: call ebx
  loc_00E04C67: mov ecx, var_18
  loc_00E04C6A: mov esi, arg_14
  loc_00E04C6D: mov ecx, [ecx+0000000Ch]
  loc_00E04C70: mov edx, [edi]
  loc_00E04C72: mov esi, [esi]
  loc_00E04C74: mov eax, [ecx+eax]
  loc_00E04C77: mov ecx, arg_1C
  loc_00E04C7A: push esi
  loc_00E04C7B: push eax
  loc_00E04C7C: push ecx
  loc_00E04C7D: push edi
  loc_00E04C7E: call [edx+00000A20h]
  loc_00E04C84: lea edx, var_18
  loc_00E04C87: push edx
  loc_00E04C88: push 00000000h
  loc_00E04C8A: call [004010E8h] ; __vbaErase
  loc_00E04C90: push 00E04CAFh
  loc_00E04C95: jmp 00E04CA2h
  loc_00E04C97: lea eax, var_1C
  loc_00E04C9A: push eax
  loc_00E04C9B: call [00401244h] ; __vbaAryUnlock
  loc_00E04CA1: ret
  loc_00E04CA2: lea ecx, var_18
  loc_00E04CA5: push ecx
  loc_00E04CA6: push 00000000h
  loc_00E04CA8: call [00401090h] ; __vbaAryDestruct
  loc_00E04CAE: ret
  loc_00E04CAF: mov ecx, var_10
  loc_00E04CB2: pop edi
  loc_00E04CB3: pop esi
  loc_00E04CB4: xor eax, eax
  loc_00E04CB6: mov fs:[00000000h], ecx
  loc_00E04CBD: pop ebx
  loc_00E04CBE: mov esp, ebp
  loc_00E04CC0: pop ebp
  loc_00E04CC1: retn 0018h
End Function

Private Sub Proc_2_157_E04CD0(arg_C) 'E04CD0
  loc_00E04CD0: mov eax, [esp+00000008h]
  loc_00E04CD4: push ebx
  loc_00E04CD5: push ebp
  loc_00E04CD6: push esi
  loc_00E04CD7: cmp eax, FFFFFFFFh
  loc_00E04CDA: push edi
  loc_00E04CDB: jnz 00E04D16h
  loc_00E04CDD: mov eax, arg_18
  loc_00E04CE1: mov ecx, arg_C
  loc_00E04CE5: push 00000000h
  loc_00E04CE7: mov [eax], 00000000h
  loc_00E04CED: mov eax, arg_20
  loc_00E04CF1: sub eax, 00000002h
  loc_00E04CF4: mov edx, [ecx]
  loc_00E04CF6: neg eax
  loc_00E04CF8: sbb eax, eax
  loc_00E04CFA: and eax, 0000002Ch
  loc_00E04CFD: add eax, 0000005Dh
  loc_00E04D00: push eax
  loc_00E04D01: mov eax, arg_28
  loc_00E04D05: push eax
  loc_00E04D06: push ecx
  loc_00E04D07: call [edx+00000A20h]
  loc_00E04D0D: pop edi
  loc_00E04D0E: pop esi
  loc_00E04D0F: pop ebp
  loc_00E04D10: xor eax, eax
  loc_00E04D12: pop ebx
  loc_00E04D13: retn 0018h
End Sub

Private Sub Proc_2_158_E04DE0(arg_C) 'E04DE0
  loc_00E04DE0: mov edx, Me
  loc_00E04DE4: mov eax, [esp+00000008h]
  loc_00E04DE8: mov ecx, Proc_2_158_E04DE0
  loc_00E04DEC: sub edx, eax
  loc_00E04DEE: jo 00E04E19h
  loc_00E04DF0: sub edx, ecx
  loc_00E04DF2: push 00000004h
  loc_00E04DF4: jo 00E04E19h
  loc_00E04DF6: sub edx, 00000004h
  loc_00E04DF9: jo 00E04E19h
  loc_00E04DFB: mov arg_C, edx
  loc_00E04DFF: lea edx, arg_C
  loc_00E04E03: add eax, ecx
  loc_00E04E05: push edx
  loc_00E04E06: jo 00E04E19h
  loc_00E04E08: push eax
  loc_00E04E09: call 006DE3F8h ; CopyMemory(%x1v, %x2v, %x3v)
  loc_00E04E0E: call [00401074h] ; __vbaSetSystemError
  loc_00E04E14: xor eax, eax
  loc_00E04E16: retn 0010h
End Sub

Private Sub Proc_2_159_E04E20
  loc_00E04E20: mov ecx, [esp+00000008h]
  loc_00E04E24: mov edx, Proc_2_159_E04E20
  loc_00E04E28: lea eax, Me
  loc_00E04E2C: add ecx, edx
  loc_00E04E2E: push 00000004h
  loc_00E04E30: push eax
  loc_00E04E31: jo 00E04E44h
  loc_00E04E33: push ecx
  loc_00E04E34: call 006DE3F8h ; CopyMemory(%x1v, %x2v, %x3v)
  loc_00E04E39: call [00401074h] ; __vbaSetSystemError
  loc_00E04E3F: xor eax, eax
  loc_00E04E41: retn 0010h
End Sub

Private Sub Proc_2_160_E04E50(arg_C) 'E04E50
  loc_00E04E50: push ebp
  loc_00E04E51: mov ebp, esp
  loc_00E04E53: sub esp, 00000008h
  loc_00E04E56: push 00402836h ; __vbaExceptHandler
  loc_00E04E5B: mov eax, fs:[00000000h]
  loc_00E04E61: push eax
  loc_00E04E62: mov fs:[00000000h], esp
  loc_00E04E69: sub esp, 00000014h
  loc_00E04E6C: push ebx
  loc_00E04E6D: push esi
  loc_00E04E6E: push edi
  loc_00E04E6F: mov var_8, esp
  loc_00E04E72: mov var_4, 00401F60h
  loc_00E04E79: mov edi, Me
  loc_00E04E7C: lea ecx, var_1C
  loc_00E04E7F: xor esi, esi
  loc_00E04E81: mov eax, [edi+0000003Ch]
  loc_00E04E84: mov var_14, esi
  loc_00E04E87: push eax
  loc_00E04E88: push ecx
  loc_00E04E89: mov var_1C, esi
  loc_00E04E8C: call [004011DCh] ; __vbaAryLock
  loc_00E04E92: mov eax, var_1C
  loc_00E04E95: cmp eax, esi
  loc_00E04E97: jz 00E04ED7h
  loc_00E04E99: cmp [eax], 0001h
  loc_00E04E9D: jnz 00E04ED7h
  loc_00E04E9F: mov ecx, arg_C
  loc_00E04EA2: mov edx, [edi]
  loc_00E04EA4: lea eax, var_14
  loc_00E04EA7: push eax
  loc_00E04EA8: push esi
  loc_00E04EA9: push ecx
  loc_00E04EAA: push edi
  loc_00E04EAB: call [edx+000009FCh]
  loc_00E04EB1: mov eax, var_1C
  loc_00E04EB4: mov esi, var_14
  loc_00E04EB7: mov edx, [eax+00000014h]
  loc_00E04EBA: mov ecx, [eax+00000010h]
  loc_00E04EBD: sub esi, edx
  loc_00E04EBF: cmp esi, ecx
  loc_00E04EC1: jb 00E04EC9h
  loc_00E04EC3: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04EC9: lea eax, [esi*8]
  loc_00E04ED0: sub eax, esi
  loc_00E04ED2: shl eax, 02h
  loc_00E04ED5: jmp 00E04EDDh
  loc_00E04ED7: call [00401100h] ; __vbaGenerateBoundsError
  loc_00E04EDD: mov edx, var_1C
  loc_00E04EE0: mov esi, [edx+0000000Ch]
  loc_00E04EE3: add esi, eax
  loc_00E04EE5: mov eax, [esi+00000008h]
  loc_00E04EE8: mov ecx, [esi]
  loc_00E04EEA: push eax
  loc_00E04EEB: push FFFFFFFCh
  loc_00E04EED: push ecx
  loc_00E04EEE: call 006DE690h ; SetWindowLong(%x1v, %x2v, %x3v)
  loc_00E04EF3: mov ebx, [00401074h] ; __vbaSetSystemError
  loc_00E04EF9: call ebx
  loc_00E04EFB: mov eax, [esi+00000004h]
  loc_00E04EFE: mov edx, [edi]
  loc_00E04F00: push 00000000h
  loc_00E04F02: push 0000005Dh
  loc_00E04F04: push eax
  loc_00E04F05: push edi
  loc_00E04F06: call [edx+00000A20h]
  loc_00E04F0C: mov edx, [esi+00000004h]
  loc_00E04F0F: mov ecx, [edi]
  loc_00E04F11: push 00000000h
  loc_00E04F13: push 00000089h
  loc_00E04F18: push edx
  loc_00E04F19: push edi
  loc_00E04F1A: call [ecx+00000A20h]
  loc_00E04F20: mov eax, [esi+00000004h]
  loc_00E04F23: push eax
  loc_00E04F24: call 006DE648h ; GlobalFree(%x1v)
  loc_00E04F29: call ebx
  loc_00E04F2B: mov edi, [004010E8h] ; __vbaErase
  loc_00E04F31: lea ecx, [esi+00000018h]
  loc_00E04F34: xor ebx, ebx
  loc_00E04F36: push ecx
  loc_00E04F37: push ebx
  loc_00E04F38: mov [esi], ebx
  loc_00E04F3A: mov [esi+0000000Ch], ebx
  loc_00E04F3D: mov [esi+00000010h], ebx
  loc_00E04F40: call edi
  loc_00E04F42: add esi, 00000014h
  loc_00E04F45: push esi
  loc_00E04F46: push ebx
  loc_00E04F47: call edi
  loc_00E04F49: lea edx, var_1C
  loc_00E04F4C: push edx
  loc_00E04F4D: call [00401244h] ; __vbaAryUnlock
  loc_00E04F53: push 00E04F63h
  loc_00E04F58: lea eax, var_1C
  loc_00E04F5B: push eax
  loc_00E04F5C: call [00401244h] ; __vbaAryUnlock
  loc_00E04F62: ret
  loc_00E04F63: mov ecx, var_10
  loc_00E04F66: pop edi
  loc_00E04F67: pop esi
  loc_00E04F68: xor eax, eax
  loc_00E04F6A: mov fs:[00000000h], ecx
  loc_00E04F71: pop ebx
  loc_00E04F72: mov esp, ebp
  loc_00E04F74: pop ebp
  loc_00E04F75: retn 0008h
End Sub

Private Sub Proc_2_161_E04F80
  loc_00E04F80: push ebx
  loc_00E04F81: mov ebx, [esp+00000008h]
  loc_00E04F85: push edi
  loc_00E04F86: mov eax, [ebx+0000003Ch]
  loc_00E04F89: push eax
  loc_00E04F8A: push 00000001h
  loc_00E04F8C: call [0040117Ch] ; __vbaUbound
  loc_00E04F92: mov edi, eax
  loc_00E04F94: test edi, edi
  loc_00E04F96: jl 00E0502Ch
  loc_00E04F9C: push ebp
  loc_00E04F9D: mov ebp, [00401100h] ; __vbaGenerateBoundsError
  loc_00E04FA3: push esi
  loc_00E04FA4: mov eax, [ebx+0000003Ch]
  loc_00E04FA7: test eax, eax
  loc_00E04FA9: jz 00E04FCFh
  loc_00E04FAB: cmp [eax], 0001h
  loc_00E04FAF: jnz 00E04FCFh
  loc_00E04FB1: mov edx, [eax+00000014h]
  loc_00E04FB4: mov ecx, [eax+00000010h]
  loc_00E04FB7: mov esi, edi
  loc_00E04FB9: sub esi, edx
  loc_00E04FBB: cmp esi, ecx
  loc_00E04FBD: jb 00E04FC1h
  loc_00E04FBF: call ebp
  loc_00E04FC1: lea eax, [esi*8]
  loc_00E04FC8: sub eax, esi
  loc_00E04FCA: shl eax, 02h
  loc_00E04FCD: jmp 00E04FD1h
  loc_00E04FCF: call ebp
  loc_00E04FD1: mov ecx, [ebx+0000003Ch]
  loc_00E04FD4: mov edx, [ecx+0000000Ch]
  loc_00E04FD7: cmp [edx+eax], 00000000h
  loc_00E04FDB: jz 00E0501Ah
  loc_00E04FDD: test ecx, ecx
  loc_00E04FDF: jz 00E05005h
  loc_00E04FE1: cmp [ecx], 0001h
  loc_00E04FE5: jnz 00E05005h
  loc_00E04FE7: mov edx, [ecx+00000014h]
  loc_00E04FEA: mov eax, [ecx+00000010h]
  loc_00E04FED: mov esi, edi
  loc_00E04FEF: sub esi, edx
  loc_00E04FF1: cmp esi, eax
  loc_00E04FF3: jb 00E04FF7h
  loc_00E04FF5: call ebp
  loc_00E04FF7: lea eax, [esi*8]
  loc_00E04FFE: sub eax, esi
  loc_00E05000: shl eax, 02h
  loc_00E05003: jmp 00E05007h
  loc_00E05005: call ebp
  loc_00E05007: mov edx, [ebx+0000003Ch]
  loc_00E0500A: mov ecx, [ebx]
  loc_00E0500C: mov edx, [edx+0000000Ch]
  loc_00E0500F: mov eax, [edx+eax]
  loc_00E05012: push eax
  loc_00E05013: push ebx
  loc_00E05014: call [ecx+00000A24h]
  loc_00E0501A: sub edi, 00000001h
  loc_00E0501D: jo 00E05033h
  loc_00E0501F: test edi, edi
  loc_00E05021: jge 00E04FA4h
  loc_00E05023: pop esi
  loc_00E05024: pop ebp
  loc_00E05025: pop edi
  loc_00E05026: xor eax, eax
  loc_00E05028: pop ebx
  loc_00E05029: retn 0004h
End Sub
