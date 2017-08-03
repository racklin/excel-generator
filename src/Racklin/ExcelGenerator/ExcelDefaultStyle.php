<?php
/**
 * Created by PhpStorm.
 * User: rack
 * Date: 2017/8/2
 * Time: 23:25
 */

namespace Racklin\ExcelGenerator;


class ExcelDefaultStyle
{

    public static $TABLE_TITLE = [
        "alignment" => [
            "horizontal" => "center",
            "vertical" => "center"
        ]
    ];

    public static $TABLE_BORDER = [
        "borders" => [
            "allborders" => [
                "style" => "thin"
            ]
        ]

    ];

    public static $HEADER = [
        "alignment" => [
            "horizontal" => "center",
            "vertical" => "center"
        ],
        "fill" => [
            "type" => "solid",
            "color" => [
                "argb" => "FFA0A0A0"
            ]
        ]
    ];

    public static $DATA = [
        "alignment" => [
            "horizontal" => "center",
            "vertical" => "center"
        ]
    ];

    public static $TITLE_ROW_HEIGHT = 30;

    public static $HEADER_ROW_HEIGHT = 20;

    public static $DATA_ROW_HEIGHT = 15;

}