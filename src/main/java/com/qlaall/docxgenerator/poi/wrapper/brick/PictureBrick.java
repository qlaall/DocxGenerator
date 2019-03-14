package com.qlaall.docxgenerator.poi.wrapper.brick;

import java.io.InputStream;
import java.util.Optional;

public class PictureBrick {
    private InputStream pictureStream;
    private int pictureType;
    private Optional<Integer> width;
    private Optional<Integer> height;

    public PictureBrick(InputStream pictureStream, int pictureType, int width, int height) {
        this.pictureStream = pictureStream;
        this.pictureType = pictureType;
        this.width = Optional.of(width);
        this.height = Optional.of(height);
    }
    public PictureBrick(InputStream pictureStream, int pictureType) {
        this.pictureStream = pictureStream;
        this.pictureType = pictureType;
        this.width = Optional.empty();
        this.height = Optional.empty();
    }
}
