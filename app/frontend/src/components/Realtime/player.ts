// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export class Player {
    private playbackNode: AudioWorkletNode | null = null;

    async init(sampleRate: number) {
        try {
            console.log("Player.init: Initializing audio context and worklet. Sample rate: ", sampleRate);
            const audioContext = new AudioContext({ sampleRate });
            console.log("Player.init: Audio context initialized. Adding worklet...");
            await audioContext.audioWorklet.addModule("playback-worklet.js");
            console.log("Player.init: Audio context and worklet initialized. Initializing playback node...");
            this.playbackNode = new AudioWorkletNode(audioContext, "playback-worklet");
            console.log("Player.init: Playback node initialized. Connecting to destination: ", audioContext.destination);
            this.playbackNode.connect(audioContext.destination);
            console.log("Player.init: Player initialized.");
        } catch (error) {
            console.error("Player.init: Error initializing audio context or worklet:", error);
        }
    }

    play(buffer: Int16Array) {
        console.log("Player.play: Posting buffer to playback node.");
        if (this.playbackNode) {
            this.playbackNode.port.postMessage(buffer);
            console.log("Player.play: Buffer posted to playback node.");
        }
    }

    clear() {
        console.log("Player.clear: Clearing playback node.");
        if (this.playbackNode) {
            this.playbackNode.port.postMessage(null);
            console.log("Player.clear: Playback node cleared.");
        }
    }
}
