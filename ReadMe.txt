PcmRiff - generate standard mono PCM WAV files in 8-bit, 16-bit and 24-bit formats at any sample rate. Note: some sound cards cannot play the 24-bit files, and many older cards can only use the 8-bit files.

--------------------------------------------------------------------------------------------

modPcmRiff.bas

PcmRiff includes a complete module for generating sine waves and modulated sine waves, in addition to mixing the waves and creating a standard Type 1 PCM WAV file. These wave files can be at any sample rate, and either 8-bit, 16-bit and 24-bit resolutions.

This code module should contain sufficient documentation and can be plugged into practically any project that requires the capability for generating complex sine wave based files. The code also includes the means for calculating attenuation functions based on a decibal (db) attenuation value.

Using an included routine "StashChunk", a complex sequence of individual waves can be strung together and compiled into a single wave. (See the embedded documentation) This feature is not demonstrated. StashChunk will save temporary PCM wave files (in an otherwise useless format) and compile them as a single (and working) file with the save function.

Added in this release: an envelope generator function. Applied against a wave form, it conforms the amplitude along the duration of the individual wave to a set of dynamic amplitude values defined by the "Envelope" subroutine.

--------------------------------------------------------------------------------------------

PCM.vbp

The project, PCM.vbp, is a program that demonstrates the use of the PcmRiff code module and includes extensive configuration options, modulation oscillation, harmonics and a string of preset "patches" for the purpose of demonstrating the capabilities of the additive sine wave approach to sound synthesis.

--------------------------------------------------------------------------------------------

PCM2.vbp

The project, PCM2.vbp, is virtually identical to the original project, but incorporates the envelope function in place of the signal level sliders. The preset "patches" are virtually identical to the original project, except that the saw tooth and square wave patches include envelopes for each wave function.

--------------------------------------------------------------------------------------------

Comments:

1. I have not finished the PhaseAngle shift in the sine wave function, nor am I certain that the feature is actually useful. The reference in the routine is both optional and inert.

--------------------------------------------------------------------------------------------

Author

All code was conceived, created and tested by the Urthman.

http://www.jsent.biz/urthman/
http://www.mp3.com/urthman/

This is a public domain package. There is no copyright (implied or otherwise) for the use of this code. Whatever you do, make it a good one, and let me know about it. A mention of credit and contribution is merely the polite thing to do.

For questions and comments, my email address can be found on either of the above URLs.