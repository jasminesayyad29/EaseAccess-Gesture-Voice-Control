import re
import time
from dataclasses import dataclass
from typing import Optional

import numpy as np
import speech_recognition as sr


COMMON_STT_FIXES = {
    "thankyou": "thank you",
    "thanks you": "thank you",
    "omegle": "omega",
    "omagle": "omega",
    "click on you": "click on two",
    "click on to": "click on two",
    "choose you": "choose two",
    "search four": "search for",
    "switch toe tab": "switch to tab",
    "switch two tab": "switch to tab",
    "go too tab": "go to tab",
    "opun": "open",
    "clik": "click",
}


@dataclass
class RecognitionConfig:
    ambient_recalibrate_every_sec: float = 90.0
    ambient_calibration_duration_sec: float = 0.18
    listen_timeout_sec: float = 1.2
    listen_phrase_time_limit_sec: float = 2.8
    min_energy_threshold: int = 180
    dynamic_energy_adjustment_damping: float = 0.15
    dynamic_energy_ratio: float = 1.55
    pause_threshold: float = 0.45
    phrase_threshold: float = 0.15
    non_speaking_duration: float = 0.2
    operation_timeout: int = 4
    enhancement_sample_rate: int = 16000
    enhancement_frame_ms: int = 20
    bandpass_low_hz: int = 85
    bandpass_high_hz: int = 3800
    noise_gate_multiplier: float = 1.65
    spectral_subtraction_strength: float = 1.15


class EnhancedVoiceRecognizer:
    def __init__(self, recognizer: Optional[sr.Recognizer] = None, config: Optional[RecognitionConfig] = None):
        self.recognizer = recognizer or sr.Recognizer()
        self.config = config or RecognitionConfig()
        self.audio_calibrated = False
        self.last_ambient_calibration_at = 0.0
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.dynamic_energy_adjustment_damping = self.config.dynamic_energy_adjustment_damping
        self.recognizer.dynamic_energy_ratio = self.config.dynamic_energy_ratio
        self.recognizer.pause_threshold = self.config.pause_threshold
        self.recognizer.phrase_threshold = self.config.phrase_threshold
        self.recognizer.non_speaking_duration = self.config.non_speaking_duration
        self.recognizer.operation_timeout = self.config.operation_timeout

    def normalize_voice_text(self, text: str) -> str:
        if not text:
            return ""
        cleaned = re.sub(r"[^a-zA-Z0-9\s]", " ", text.lower())
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        for wrong, right in COMMON_STT_FIXES.items():
            cleaned = re.sub(rf"\b{re.escape(wrong)}\b", right, cleaned)
        return cleaned

    def _enhance_audio_for_recognition(self, audio: sr.AudioData):
        if audio is None:
            return None

        try:
            raw_data = audio.get_raw_data(convert_rate=self.config.enhancement_sample_rate, convert_width=2)
            samples = np.frombuffer(raw_data, dtype=np.int16).astype(np.float32)
            if samples.size < (self.config.enhancement_sample_rate // 8):
                return audio

            samples /= 32768.0
            samples -= float(np.mean(samples))

            pre_emphasis = np.empty_like(samples)
            pre_emphasis[0] = samples[0]
            pre_emphasis[1:] = samples[1:] - 0.97 * samples[:-1]

            frame_size = max(1, int(self.config.enhancement_sample_rate * self.config.enhancement_frame_ms / 1000))
            pad_size = (-len(pre_emphasis)) % frame_size
            if pad_size:
                pre_emphasis = np.pad(pre_emphasis, (0, pad_size), mode="constant")

            frames = pre_emphasis.reshape(-1, frame_size)
            frame_energy = np.sqrt(np.mean(frames ** 2, axis=1) + 1e-12)
            if frame_energy.size:
                quiet_count = max(1, int(len(frame_energy) * 0.25))
                quiet_frames = np.partition(frame_energy, quiet_count - 1)[:quiet_count]
                noise_floor = float(np.median(quiet_frames))
            else:
                noise_floor = 0.0

            spectrum = np.fft.rfft(pre_emphasis)
            freqs = np.fft.rfftfreq(len(pre_emphasis), d=1.0 / self.config.enhancement_sample_rate)
            band_mask = (freqs >= self.config.bandpass_low_hz) & (freqs <= self.config.bandpass_high_hz)
            spectrum *= band_mask
            cleaned = np.fft.irfft(spectrum, n=len(pre_emphasis))

            if noise_floor > 1e-6:
                gate_threshold = max(noise_floor * self.config.noise_gate_multiplier, 0.006)
                envelope = np.clip((np.abs(cleaned) - gate_threshold) / (1.0 - gate_threshold), 0.0, 1.0)
                cleaned = np.sign(cleaned) * np.abs(cleaned) * (0.35 + (0.65 * envelope))
                noise_suppression = np.clip(
                    1.0 - (noise_floor * self.config.spectral_subtraction_strength),
                    0.25,
                    1.0,
                )
                cleaned *= noise_suppression

            peak = float(np.max(np.abs(cleaned))) if cleaned.size else 0.0
            if peak > 0:
                cleaned = cleaned / peak * 0.92

            enhanced = np.clip(cleaned, -1.0, 1.0)
            enhanced_raw = (enhanced * 32767.0).astype(np.int16).tobytes()
            return sr.AudioData(enhanced_raw, self.config.enhancement_sample_rate, 2)
        except Exception:
            return audio

    def _score_command_candidate(self, candidate, command_library):
        transcript = self.normalize_voice_text(candidate.get("transcript", ""))
        if not transcript:
            return -1.0, ""

        stt_confidence = float(candidate.get("confidence", 0.0))
        similarity = 0.0
        if command_library:
            from difflib import SequenceMatcher
            similarity = max(SequenceMatcher(None, transcript, known).ratio() for known in command_library)

        score = (0.65 * similarity) + (0.35 * stt_confidence)
        return score, transcript

    def best_transcript_from_google(self, audio: sr.AudioData, command_library):
        primary = self.normalize_voice_text(self.recognizer.recognize_google(audio))
        if primary:
            if len(primary.split()) >= 2:
                return primary
            if any(k in primary for k in ("omega", "open", "close", "click", "scroll", "search", "tab", "choose")):
                return primary

        enhanced_audio = self._enhance_audio_for_recognition(audio)
        if enhanced_audio is not None and enhanced_audio is not audio:
            try:
                primary = self.normalize_voice_text(self.recognizer.recognize_google(enhanced_audio))
                if primary:
                    if len(primary.split()) >= 2:
                        return primary
                    if any(k in primary for k in ("omega", "open", "close", "click", "scroll", "search", "tab", "choose")):
                        return primary
            except sr.UnknownValueError:
                pass

        for candidate_audio in (enhanced_audio, audio):
            if candidate_audio is None:
                continue
            try:
                alt_data = self.recognizer.recognize_google(candidate_audio, show_all=True)
            except sr.UnknownValueError:
                continue

            if isinstance(alt_data, dict):
                alternatives = alt_data.get("alternative", [])
                if alternatives:
                    ranked = [self._score_command_candidate(candidate, command_library) for candidate in alternatives[:6]]
                    ranked = [item for item in ranked if item[1]]
                    if ranked:
                        ranked.sort(key=lambda item: item[0], reverse=True)
                        return ranked[0][1]

        return primary

    def capture_and_recognize(self, mic_source, command_library):
        now = time.time()
        should_recalibrate = (
            (not self.audio_calibrated)
            or (now - self.last_ambient_calibration_at) > self.config.ambient_recalibrate_every_sec
        )
        if should_recalibrate:
            self.recognizer.adjust_for_ambient_noise(mic_source, duration=self.config.ambient_calibration_duration_sec)
            self.audio_calibrated = True
            self.last_ambient_calibration_at = now

        self.recognizer.energy_threshold = max(self.config.min_energy_threshold, int(self.recognizer.energy_threshold))
        audio = self.recognizer.listen(
            mic_source,
            timeout=self.config.listen_timeout_sec,
            phrase_time_limit=self.config.listen_phrase_time_limit_sec,
        )
        voice_data = self.best_transcript_from_google(audio, command_library)
        return self.normalize_voice_text(voice_data)
