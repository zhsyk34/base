package com.cat.audio;

/**
 * Created by Administrator on 16-5-26.
 */
public class Test {

    public static void main(String[] args) throws Exception {
        pre();
        // play();
    }

    public static void pre() throws Exception {
        String name = "mp3/1.mp3";
        String path = Util.getPath("") + name;
        System.out.println(path);

        AudioConvert.convertToWav(path, null);
    }

    public static void play() {
        String path = "mp3/1.wav";
        AudioPlayer.beep(path);
    }
}
