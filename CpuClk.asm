.586                                        ;# Create 32 bit code, pentium instructions allowed
.model flat, stdcall                        ;# 32 bit memory model
option casemap :none                        ;# Case sensitive

.code
start:
CpuClk proc
    rdtsc                                   ;# Read the cpu clock cycle count into eax/edx
    mov ecx, dword ptr [esp+08h]            ;# Address of Currency parameter into ecx
    mov [ecx], eax                          ;# Put eax to the low part of the currency variable
    mov [ecx+4], edx                        ;# Put edx to the high part of the currency variable
    xor eax, eax                            ;# Clear eax
    ret 8                                   ;# Return
CpuClk endp
end start
