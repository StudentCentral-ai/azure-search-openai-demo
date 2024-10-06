class PlaybackWorklet extends AudioWorkletProcessor {
    constructor() {
        console.log("PlaybackWorklet: constructor");
        super();
        this.port.onmessage = this.handleMessage.bind(this);
        this.port.on;
        this.buffer = [];
        console.log("PlaybackWorklet: constructor ended");
    }

    handleMessage(event) {
        console.log("PlaybackWorklet.handleMessage: ", event.data);
        if (event.data === null) {
            this.buffer = [];
            return;
        }
        this.buffer.push(...event.data);
        console.log("PlaybackWorklet.handleMessage: ended");
    }

    process(inputs, outputs, parameters) {
        //console.log("PlaybackWorklet.process");
        const output = outputs[0];
        const channel = output[0];

        if (this.buffer.length > channel.length) {
            const toProcess = this.buffer.slice(0, channel.length);
            this.buffer = this.buffer.slice(channel.length);
            channel.set(toProcess.map(v => v / 32768));
        } else {
            channel.set(this.buffer.map(v => v / 32768));
            this.buffer = [];
        }
        //console.log("PlaybackWorklet.process ended");
        return true;
    }
}

registerProcessor("playback-worklet", PlaybackWorklet);
