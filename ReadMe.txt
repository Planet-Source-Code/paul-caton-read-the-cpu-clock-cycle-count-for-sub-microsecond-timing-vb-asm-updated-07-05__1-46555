Read the cpu clock cycle count for sub-microsecond timing. VB+ASM

Pentium class cpu's include a 64 bit register that increments from power-on at the CPU clock frequency. With a 2GHz processor, you have in effect, a 2GHz clock. The cCpuClk class allows the user to retrieve the 64 bit CPU clock cycle count into a passed currency parameter. The class can be used as a basis for sub-microsecond benchmarking and delay timing. Note that with the extreme resolution provided, multitasking and the state of the cpu caches will show in the results. Thanks to David Fritts and Robert Rayment for the vtable trick.

Paul Caton
Paul_Caton@hotmail.com