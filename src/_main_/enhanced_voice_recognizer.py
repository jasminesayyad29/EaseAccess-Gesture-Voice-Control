import json
import os
import re
import time
from dataclasses import dataclass
from typing import List, Optional

import numpy as np
import requests
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
    "opened": "open",
    "opening": "open",
    "launched": "launch",
    "launching": "launch",
    "started": "start",
    "starting": "start",
    "ran": "run",
    "running": "run",
    "closed": "close",
    "closing": "close",
    "quit application": "close application",
    "quitted": "quit",
    "exited": "exit",
    "stopped": "stop",
    "ending": "end",
    "terminated": "terminate",
    "killed": "kill",
    "clicked": "click",
    "clicking": "click",
    "tapped": "tap",
    "tapping": "tap",
    "pressed": "press",
    "pressing": "press",
    "searched": "search",
    "searching": "search",
    "found": "find",
    "looking up": "look up",
    "selected": "select",
    "selecting": "select",
    "chose": "choose",
    "chosen": "choose",
    "choosing": "choose",
    "copied": "copy",
    "copying": "copy",
    "pasted": "paste",
    "pasting": "paste",
    "typed": "type",
    "typing": "type",
    "wrote": "write",
    "writing": "write",
    "scrolled": "scroll",
    "scrolling": "scroll",
    "maximized": "maximize",
    "maximised": "maximize",
    "maximizing": "maximize",
    "minimized": "minimize",
    "minimised": "minimize",
    "minimizing": "minimize",
}

PLANNED_COMMAND_SEPARATOR = " || "
GEMINI_MIN_WORDS = 9
GEMINI_INTENT_MODELS = (
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash",
    "gemini-1.5-flash-latest",
)

SUPPORTED_GEMINI_INTENTS = {
    "open_app",
    "search",
    "location",
    "click_action",
    "choice_select",
    "tab_control",
    "copy",
    "paste",
    "range_select",
    "presentation_control",
    "time_query",
    "date_query",
    "name_query",
    "greeting",
    "goodbye",
}


@dataclass
class RecognitionConfig:
    ambient_recalibrate_every_sec: float = 90.0
    ambient_calibration_duration_sec: float = 0.12
    listen_timeout_sec: float = 1.5
    listen_phrase_time_limit_sec: float = 7.0
    min_energy_threshold: int = 180
    dynamic_energy_adjustment_damping: float = 0.15
    dynamic_energy_ratio: float = 1.45
    pause_threshold: float = 0.72
    phrase_threshold: float = 0.15
    non_speaking_duration: float = 0.32
    operation_timeout: int = 6
    enhancement_sample_rate: int = 16000
    enhancement_frame_ms: int = 20
    bandpass_low_hz: int = 85
    bandpass_high_hz: int = 3800
    noise_gate_multiplier: float = 1.45
    spectral_subtraction_strength: float = 0.85
    trim_silence_padding_ms: int = 140
    min_fast_confidence: float = 0.62


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
        self._load_local_env()
        self._gemini_disabled_until = 0.0
        self._gemini_model_cache = {"models": None, "at": 0.0}

    def normalize_voice_text(self, text: str) -> str:
        if not text:
            return ""
        cleaned = re.sub(r"[^a-zA-Z0-9\s]", " ", text.lower())
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        for wrong, right in COMMON_STT_FIXES.items():
            cleaned = re.sub(rf"\b{re.escape(wrong)}\b", right, cleaned)
        return cleaned

    def _load_local_env(self):
        candidate_paths = [
            os.path.join(os.path.dirname(__file__), ".env"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), ".env"),
            os.path.join(os.getcwd(), ".env"),
        ]
        env_path = next((p for p in candidate_paths if os.path.exists(p)), None)
        if not env_path:
            return
        try:
            with open(env_path, "r", encoding="utf-8") as env_file:
                for raw_line in env_file:
                    line = raw_line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    if key and value and key not in os.environ:
                        os.environ[key] = value
        except Exception:
            pass

    def _strip_wake_word(self, text: str) -> str:
        if not text:
            return ""
        return re.sub(r"^\s*(omega|omegaa|omeg|ome|om|oh mega|o mega|mega)\s+", "", text).strip()

    def _ensure_omega_command(self, command: str) -> str:
        cleaned = self.normalize_voice_text(command)
        if not cleaned:
            return ""
        if not cleaned.startswith("omega "):
            cleaned = f"omega {cleaned}"
        return cleaned

    def _command_from_action(self, action: dict) -> str:
        if not isinstance(action, dict):
            return ""

        command = action.get("command")
        if isinstance(command, str) and command.strip():
            return self._ensure_omega_command(command)

        intent = str(action.get("intent", "")).strip().lower()
        entity = self.normalize_voice_text(str(action.get("entity", "")).strip())
        if intent not in SUPPORTED_GEMINI_INTENTS:
            return ""

        if intent == "open_app" and entity:
            return self._ensure_omega_command(f"open {entity}")
        if intent == "search" and entity:
            return self._ensure_omega_command(f"search for {entity}")
        if intent == "location" and entity:
            return self._ensure_omega_command(f"where is {entity}")
        if intent == "click_action":
            return self._ensure_omega_command(f"click {entity}".strip())
        if intent == "choice_select":
            if not entity:
                return ""
            if entity.startswith(("choose ", "select ", "option ")):
                return self._ensure_omega_command(entity)
            return self._ensure_omega_command(f"choose {entity}")
        if intent == "tab_control":
            return self._ensure_omega_command(entity or "next tab")
        if intent == "presentation_control":
            return self._ensure_omega_command(entity)
        if intent in {"copy", "paste", "range_select", "time_query", "date_query", "name_query", "greeting", "goodbye"}:
            return self._ensure_omega_command(entity or intent.replace("_", " "))
        return ""

    def _extract_json_object(self, text: str):
        if not text:
            return None
        cleaned = text.strip()
        cleaned = re.sub(r"^```(?:json)?", "", cleaned, flags=re.IGNORECASE).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()
        try:
            return json.loads(cleaned)
        except Exception:
            match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
            if not match:
                return None
            try:
                return json.loads(match.group(0))
            except Exception:
                return None

    def _discover_gemini_generate_models(self, api_key: str) -> List[str]:
        now = time.time()
        cached_models = self._gemini_model_cache.get("models")
        if cached_models and (now - self._gemini_model_cache.get("at", 0.0)) < 600:
            return cached_models

        discovered = []
        for api_ver in ("v1beta", "v1"):
            url = f"https://generativelanguage.googleapis.com/{api_ver}/models?key={api_key}"
            try:
                resp = requests.get(url, timeout=8)
                data = resp.json() if resp.content else {}
                if not resp.ok:
                    continue
                for item in data.get("models", []):
                    methods = item.get("supportedGenerationMethods", []) or []
                    if "generateContent" not in methods:
                        continue
                    name = item.get("name", "")
                    if name.startswith("models/"):
                        name = name.split("/", 1)[1]
                    if name:
                        discovered.append(name)
            except Exception:
                continue

        def sort_key(model_name):
            lowered = model_name.lower()
            if "flash" in lowered and "latest" not in lowered:
                rank = 0
            elif "flash" in lowered:
                rank = 1
            else:
                rank = 2
            return rank, model_name

        ordered = sorted(set(discovered), key=sort_key)
        fallback = [name for name in GEMINI_INTENT_MODELS if name not in ordered]
        models = ordered + fallback
        self._gemini_model_cache = {"models": models, "at": now}
        return models

    def _gemini_intent_plan(self, transcript: str, command_library=None) -> List[str]:
        api_key = (
            os.getenv("GEMINI_API_KEY", "").strip()
            or os.getenv("GOOGLE_API_KEY", "").strip()
        )
        if not api_key or not transcript:
            return []
        if time.time() < self._gemini_disabled_until:
            return []

        examples = list(command_library or [])[:60]
        prompt = (
            "You are the intent planner for a Windows voice assistant named Omega.\n"
            "Convert the user's recognized speech into one or more executable Omega commands.\n"
            "Return ONLY valid JSON with this exact shape:\n"
            '{"actions":[{"intent":"open_app","action":"open","entity":"chrome","command":"omega open chrome"}]}\n'
            "Supported intents: open_app, search, location, click_action, choice_select, tab_control, copy, paste, "
            "range_select, presentation_control, time_query, date_query, name_query, greeting, goodbye.\n"
            "Rules:\n"
            "- Split compound requests into ordered actions.\n"
            "- Keep commands short and directly executable by the assistant.\n"
            "- Preserve search queries exactly except for filler words.\n"
            "- Do not invent actions that were not requested.\n"
            "- If the input is already a simple command, return one action.\n"
            "Example input: omega please open chrome for me and search for image processing\n"
            "Example output: "
            '{"actions":[{"intent":"open_app","action":"open","entity":"chrome","command":"omega open chrome"},'
            '{"intent":"search","action":"search","entity":"image processing","command":"omega search for image processing"}]}\n'
            "Example input: omega please open wps and click on Module and select option 5\n"
            "Example output: "
            '{"actions":[{"intent":"open_app","action":"open","entity":"wps","command":"omega open wps"},'
            '{"intent":"click_action","action":"click","entity":"module","command":"omega click module"},'
            '{"intent":"choice_select","action":"choose","entity":"5","command":"omega choose 5"}]}\n'
            f"Known command examples: {examples}\n"
            f"User input: {transcript}"
        )
        base_payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.1,
                "maxOutputTokens": 512,
            },
        }
        json_mode_payload = json.loads(json.dumps(base_payload))
        json_mode_payload["generationConfig"]["responseMimeType"] = "application/json"
        candidate_models = self._discover_gemini_generate_models(api_key)

        for api_ver in ("v1beta", "v1"):
            for model_name in candidate_models:
                url = f"https://generativelanguage.googleapis.com/{api_ver}/models/{model_name}:generateContent?key={api_key}"
                for payload in (json_mode_payload, base_payload):
                    try:
                        resp = requests.post(url, json=payload, timeout=8)
                        data = resp.json() if resp.content else {}
                        if not resp.ok:
                            continue
                        parts = data.get("candidates", [{}])[0].get("content", {}).get("parts", [])
                        text = " ".join(
                            part.get("text", "") for part in parts if isinstance(part, dict)
                        ).strip()
                        parsed = self._extract_json_object(text)
                        actions = parsed.get("actions", []) if isinstance(parsed, dict) else []
                        commands = [self._command_from_action(action) for action in actions]
                        commands = [cmd for cmd in commands if cmd]
                        if commands:
                            return commands[:5]
                    except Exception:
                        continue
        self._gemini_disabled_until = time.time() + 120.0
        return []

    def _fallback_intent_plan(self, transcript: str) -> List[str]:
        stripped = self._strip_wake_word(self.normalize_voice_text(transcript))
        if not stripped:
            return []

        commands = []
        clauses = [
            clause.strip()
            for clause in re.split(r"\b(?:and then|then|and|also)\b", stripped)
            if clause.strip()
        ]
        for clause in clauses:
            open_match = re.search(
                r"\b(?:open|launch|start|run)(?:\s+(?:the|app|application|program))?\s+(.+)$",
                clause,
            )
            search_match = re.search(r"\b(?:search|find|look up|google)(?:\s+(?:for|about))?\s+(.+)$", clause)
            click_match = re.search(r"\b(?:click on|click|tap|press|select)\s+(.+)$", clause)
            choice_match = re.search(r"\b(?:choose|select|option)\s+(?:option\s+)?([a-z0-9]+)\b", clause)

            if open_match:
                app = re.sub(r"\b(for me|please|kindly|now|the)\b", " ", open_match.group(1))
                app = re.sub(r"\s+", " ", app).strip()
                if app:
                    commands.append(self._ensure_omega_command(f"open {app}"))
                    continue

            if search_match:
                query = re.sub(r"\b(please|for me|kindly|now)\b", " ", search_match.group(1))
                query = re.sub(r"\s+", " ", query).strip()
                if query:
                    commands.append(self._ensure_omega_command(f"search for {query}"))
                    continue

            if choice_match:
                commands.append(self._ensure_omega_command(f"choose {choice_match.group(1)}"))
                continue

            if click_match:
                target = re.sub(r"\b(please|for me|kindly|now|the)\b", " ", click_match.group(1))
                target = re.sub(r"\boption\s+[a-z0-9]+\b", " ", target)
                target = re.sub(r"\s+", " ", target).strip()
                if target:
                    commands.append(self._ensure_omega_command(f"click {target}"))
                    continue

        if len(commands) > 1:
            return commands[:5]

        search_match = re.search(r"\b(?:search|find|look up|google)(?:\s+(?:for|about))?\s+(.+)$", stripped)
        open_part = re.split(
            r"\b(?:and then|then|and|also)\s+(?:search|find|look up|google)\b|\b(?:search|find|look up|google)\b",
            stripped,
            maxsplit=1,
        )[0]
        open_match = re.search(r"\b(?:open|launch|start|run)(?:\s+(?:the|app|application|program))?\s+(.+)$", open_part)

        if open_match:
            app = open_match.group(1)
            app = re.sub(r"\b(for me|please|kindly|now|the)\b", " ", app)
            app = re.sub(r"\s+", " ", app).strip()
            if app:
                commands.append(self._ensure_omega_command(f"open {app}"))

        if search_match:
            query = search_match.group(1)
            query = re.sub(r"\b(please|for me|kindly|now)\b", " ", query)
            query = re.sub(r"\s+", " ", query).strip()
            if query:
                commands.append(self._ensure_omega_command(f"search for {query}"))

        if len(commands) > 1:
            return commands
        return []

    def plan_commands_with_gemini(self, transcript: str, command_library=None) -> List[str]:
        normalized = self.normalize_voice_text(transcript)
        if not normalized:
            return []
        if len(normalized.split()) < GEMINI_MIN_WORDS:
            return []

        planned = self._gemini_intent_plan(normalized, command_library)
        if planned:
            return planned
        return self._fallback_intent_plan(normalized)

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

            frame_size = max(1, int(self.config.enhancement_sample_rate * self.config.enhancement_frame_ms / 1000))
            pad_size = (-len(samples)) % frame_size
            framed_samples = np.pad(samples, (0, pad_size), mode="constant") if pad_size else samples
            frames_for_trim = framed_samples.reshape(-1, frame_size)
            frame_energy_for_trim = np.sqrt(np.mean(frames_for_trim ** 2, axis=1) + 1e-12)
            if frame_energy_for_trim.size:
                floor = float(np.percentile(frame_energy_for_trim, 25))
                trim_threshold = max(floor * 2.2, 0.008)
                voiced = np.where(frame_energy_for_trim > trim_threshold)[0]
                if voiced.size:
                    pad_frames = max(1, int(self.config.trim_silence_padding_ms / self.config.enhancement_frame_ms))
                    start_frame = max(0, int(voiced[0]) - pad_frames)
                    end_frame = min(len(frames_for_trim), int(voiced[-1]) + pad_frames + 1)
                    samples = framed_samples[start_frame * frame_size:end_frame * frame_size]
                    if samples.size < (self.config.enhancement_sample_rate // 8):
                        return audio

            pre_emphasis = np.empty_like(samples)
            pre_emphasis[0] = samples[0]
            pre_emphasis[1:] = samples[1:] - 0.97 * samples[:-1]

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

    def _is_command_like(self, transcript: str) -> bool:
        return any(
            keyword in transcript
            for keyword in (
                "omega", "open", "close", "click", "scroll", "search", "tab", "choose",
                "select", "copy", "paste", "type", "write", "next", "back",
            )
        )

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
        primary = ""
        try:
            alt_data = self.recognizer.recognize_google(audio, show_all=True)
        except sr.UnknownValueError:
            alt_data = None

        if isinstance(alt_data, dict):
            alternatives = alt_data.get("alternative", [])
            if alternatives:
                top = alternatives[0]
                primary = self.normalize_voice_text(top.get("transcript", ""))
                confidence = float(top.get("confidence", 0.0))
                if primary and (
                    confidence >= self.config.min_fast_confidence
                    or len(primary.split()) >= 4
                    or self._is_command_like(primary)
                ):
                    return primary

                ranked = [self._score_command_candidate(candidate, command_library) for candidate in alternatives[:5]]
                ranked = [item for item in ranked if item[1]]
                if ranked:
                    ranked.sort(key=lambda item: item[0], reverse=True)
                    best_score, best_transcript = ranked[0]
                    if best_score >= 0.35 or self._is_command_like(best_transcript):
                        return best_transcript
        else:
            try:
                primary = self.normalize_voice_text(self.recognizer.recognize_google(audio))
                if primary and (len(primary.split()) >= 4 or self._is_command_like(primary)):
                    return primary
            except sr.UnknownValueError:
                primary = ""

        enhanced_audio = self._enhance_audio_for_recognition(audio)
        if enhanced_audio is not None and enhanced_audio is not audio:
            try:
                enhanced_data = self.recognizer.recognize_google(enhanced_audio, show_all=True)
                if isinstance(enhanced_data, dict):
                    alternatives = enhanced_data.get("alternative", [])
                    ranked = [self._score_command_candidate(candidate, command_library) for candidate in alternatives[:6]]
                    ranked = [item for item in ranked if item[1]]
                    if ranked:
                        ranked.sort(key=lambda item: item[0], reverse=True)
                        return ranked[0][1]
                enhanced_primary = self.normalize_voice_text(self.recognizer.recognize_google(enhanced_audio))
                if enhanced_primary:
                    return enhanced_primary
            except sr.UnknownValueError:
                pass

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
        normalized_voice = self.normalize_voice_text(voice_data)
        planned_commands = self.plan_commands_with_gemini(normalized_voice, command_library)
        if planned_commands:
            return PLANNED_COMMAND_SEPARATOR.join(planned_commands)
        return normalized_voice
