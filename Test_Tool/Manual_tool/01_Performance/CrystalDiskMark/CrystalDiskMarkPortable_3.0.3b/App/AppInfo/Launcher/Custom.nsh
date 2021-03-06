${SegmentFile}

${Segment.OnInit}
	; Borrowed the following from PAL 2.2, Remove on release of PAL 2.2
		; Work out if it's 64-bit or 32-bit
	System::Call kernel32::GetCurrentProcess()i.s
	System::Call kernel32::IsWow64Process(is,*i.r0)
	${If} $0 == 0
		;32
		Rename "$EXEDIR\Data\DiskMarkX64.ini" "$EXEDIR\Data\DiskMark.ini"
	${Else}
		;64
		Rename "$EXEDIR\Data\DiskMark.ini" "$EXEDIR\Data\DiskMarkX64.ini"
	${EndIf}
!macroend
