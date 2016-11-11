<?php
// with Composer
require_once 'vendor/autoload.php';

use AlejandroSosa\YiiPowerPoint\PowerPoint;

//setting file
$options = [
    //file
    'fileName' => Yii::t('app', 'demo'),

    //properties of file
    'fileProperties' => [
        'creator' => Yii::t('app', 'MyCompany'),
        'title' => Yii::t('app', 'Financial Statement 2016'),
        'subject' => Yii::t('app', 'General resume'),
        'description' => Yii::t('app', 'Budget report')
    ],

    //design all ppt
    'layout' => [
        'background' => Yii::getPathOfAlias('images') .'/bg.png'
    ],
];

//slides of ppt
$slides = [
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'My PPT'),
                'options'=>[
                    'height'=>82, 'width'=>816, 'ox'=>72, 'oy'=>170,
                    'bold'=> true, 'size'=>44, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', 'My custom ppt example 2016'),
                'options'=>[
                    'height'=>82, 'width'=>816,
                    'ox'=>72, 'oy'=>291,
                    'bold'=> false, 'size'=>40, 'color'=>PowerPoint::COLOR_PRIMARY_TEXT,
                    'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER,
                ],
            ],
        ]
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Available options'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>52,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', 'Table of Contents'),
                'options'=>[
                    'height'=>38, 'width'=>413, 'ox'=>50, 'oy'=>110,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_PRIMARY_TEXT,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', 'Page'),
                'options'=>[
                    'height'=>38, 'width'=>91, 'ox'=>706, 'oy'=>110,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_PRIMARY_TEXT,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', '3{break}4{break}5{break}6{break}7{break}8{break}9{break}10',
                    ['{break}'=>PowerPoint::TEXT_BREAK]),
                'options'=>[
                    'height'=>220, 'width'=>91, 'ox'=>706, 'oy'=>163,
                    'bold'=> false, 'size'=>14, 'color'=>PowerPoint::COLOR_PRIMARY_TEXT,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app','1. Tables {break}2. Chart bar {break}3. Chart bar horizontal {break}4. Chart bar horizontal percent stacked {break}5. Chart bar 3D {break}6. Chart scatter {break}7. Chart pie {break}8. Chart pie 3D',
                    ['{break}'=>PowerPoint::TEXT_BREAK]),
                'options'=>[
                    'height'=>220, 'width'=>604, 'ox'=>102, 'oy'=>163,
                    'bold'=>false, 'size'=>14, 'color'=>PowerPoint::COLOR_PRIMARY_TEXT,
                    'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_LEFT,
                ]
            ],
        ],
        'images' => [
            'path'=>Yii::getPathOfAlias('images') . '/ppt/pptflujo.png',
            'height'=>143, 'width'=>793, 'ox'=>120, 'oy'=>180,
            'resizeProportional'=> false,
            'name'=>Yii::t('app', 'Name of image'),
            'description'=>Yii::t('app', 'This is a demo description')
        ]
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Table'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ]
            ]
        ],
        'tables' => [
            [
                'height'=>400, 'width'=>900, 'ox'=>26, 'oy'=>163,
                'header' => [
                    'columns'=> [
                        ['text'=>'Main Title', 'style'=>['width'=>268]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                    ],
                    'style'=>[
                        'background'=>'FFC6D9F1', 'bold'=> true, 'size'=>9, 'color'=>PowerPoint::COLOR_PRIMARY_TEXT,
                        'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER, 'height'=>35
                    ],
                ],
                'rows'=>[
                    [
                        'columns'=>[
                            //with style
                            ['text'=>'Title of this row', 'style'=>['align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_LEFT]],
                            //without style
                            '2.140',
                            '26,07%',
                            '9.00',
                            '2.121',
                            '1.554',
                            '7.000',
                            '73,27%'
                        ],
                        'style'=>[
                            'background'=>'FFD7E4BD', 'bold'=> true, 'size'=>10, 'color'=>PowerPoint::COLOR_PRIMARY_TEXT,
                            'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_RIGHT, 'height'=>17
                        ]
                    ],
                    [
                        'columns'=>[
                            //with style
                            ['text'=>'Second title of this row', 'style'=>['align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_LEFT]],
                            //without style
                            '2.140',
                            '26,07%',
                            '9.00',
                            '2.121',
                            '1.554',
                            '7.000',
                            '73,27%'
                        ],
                        'style'=>[
                            'background'=>'FFE9EDF4', 'bold'=> false, 'size'=>10, 'color'=>PowerPoint::COLOR_PRIMARY_TEXT,
                            'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_RIGHT, 'height'=>17
                        ]
                    ],
                ],
            ],
        ],
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Chart Bar'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ]
        ],
        'charts' => [
            [
                'title'=>Yii::t('app', 'My first chart bar'),
                'series'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 7, 'Tuesday' => 9, 'Wednesday' => 19, 'Thursday' => 16, 'Friday' => 8, 'Saturday' => 11],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR, 'seriesName'=> ['2014', '2015', '2016'], 'showSeriesName'=> false,
                    'legendItalic' => true, 'legendOx'=> 20,'legendOy'=> 20,
                    'titleItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My second chart bar'),
                'series'=> ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR, 'seriesName'=> ['2009', '2010'], 'showSeriesName'=> false,
                    'legendItalic' => true, 'legendOx'=> 20,'legendOy'=> 20, 'legendBackground'=>PowerPoint::COLOR_AMBER,
                    'titleItalic' => true, 'background' => PowerPoint::COLOR_BLUE,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My third chart bar'),
                'series'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR, 'seriesName'=> ['2009', '2010'], 'showSeriesName'=> false,
                    'titleItalic' => true, 'legendItalic' => true, 'legendVisible' => false,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>367
                ]
            ],
            [
                'title'=>Yii::t('app', 'My quarter chart bar'),
                'series'=>
                    [
                        ['October' => 12, 'November' => 15, 'December' => 13],
                        ['October' => 11, 'November' => 20, 'December' => 14]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR, 'seriesName'=> ['2014', '2015'], 'showSeriesName'=> false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>367
                ]
            ]
        ],
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Chart bar horizontal'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ]
        ],
        'charts' => [
            [
                'title'=>Yii::t('app', 'My first chart bar horizontal'),
                'series'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 7, 'Tuesday' => 9, 'Wednesday' => 19, 'Thursday' => 16, 'Friday' => 8, 'Saturday' => 11],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_HORIZONTAL, 'seriesName'=> ['2014', '2015', '2016'], 'showSeriesName'=> false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My second chart bar horizontal'),
                'series'=> ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_HORIZONTAL, 'seriesName'=> ['2009', '2010'], 'showSeriesName'=> false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My third chart bar horizontal'),
                'series'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_HORIZONTAL, 'seriesName'=> ['2009', '2010'], 'showSeriesName'=> false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>367
                ]
            ],
            [
                'title'=>Yii::t('app', 'My quarter chart bar horizontal'),
                'series'=>
                    [
                        ['October' => 12, 'November' => 15, 'December' => 13],
                        ['October' => 11, 'November' => 20, 'December' => 14]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_HORIZONTAL, 'seriesName'=> ['2014', '2015'], 'showSeriesName'=> false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>367
                ]
            ]
        ],
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Chart bar horizontal percent stacked'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ]
        ],
        'charts' => [
            [
                'title'=>Yii::t('app', 'My first chart bar horizontal percent stacked'),
                'series'=>
                    [
                        ['Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293],
                        ['Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379],
                        ['Jan' => 233, 'Feb' => 146, 'Mar' => 238, 'Apr' => 175, 'May' => 108, 'Jun' => 257, 'Jul' => 199, 'Aug' => 201, 'Sep' => 88, 'Oct' => 147, 'Nov' => 287, 'Dec' => 105]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_PERCENT_STACKED, 'seriesName'=> ['2015', '2016', '2017'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My second chart bar horizontal percent stacked'),
                'series'=> ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_PERCENT_STACKED, 'seriesName'=> ['2015'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My third chart bar horizontal percent stacked'),
                'series'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_PERCENT_STACKED, 'seriesName'=> ['2015', '2016'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>367
                ]
            ],
            [
                'title'=>Yii::t('app', 'My quarter chart bar horizontal percent stacked'),
                'series'=>
                    [
                        ['October' => 12, 'November' => 15, 'December' => 13],
                        ['October' => 11, 'November' => 20, 'December' => 14]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_PERCENT_STACKED, 'seriesName'=> ['2015', '2016'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>367
                ]
            ]
        ],
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Chart bar 3D'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ]
        ],
        'charts' => [
            [
                'title'=>Yii::t('app', 'My first chart bar 3D'),
                'series'=>
                    [
                        ['Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293],
                        ['Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379],
                        ['Jan' => 233, 'Feb' => 146, 'Mar' => 238, 'Apr' => 175, 'May' => 108, 'Jun' => 257, 'Jul' => 199, 'Aug' => 201, 'Sep' => 88, 'Oct' => 147, 'Nov' => 287, 'Dec' => 105]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_3D, 'seriesName'=> ['2015', '2016', '2017'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My second chart bar 3D'),
                'series'=> ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_3D, 'seriesName'=> ['2015'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My third chart bar 3D'),
                'series'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_3D, 'seriesName'=> ['2015', '2016'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>367
                ]
            ],
            [
                'title'=>Yii::t('app', 'My quarter chart bar 3D'),
                'series'=>
                    [
                        ['October' => 12, 'November' => 15, 'December' => 13],
                        ['October' => 11, 'November' => 20, 'December' => 14]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_3D, 'seriesName'=> ['2015', '2016'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>367
                ]
            ]
        ],
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Chart scatter'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ]
        ],
        'charts' => [
            [
                'title'=>Yii::t('app', 'My first chart scatter'),
                'series'=> ['Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_SCATTER, 'seriesName'=> ['2015'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My second chart scatter'),
                'series'=> ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_SCATTER, 'seriesName'=> ['2015'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'showSymbol' => true, 'typeSymbol' => PowerPoint::SYMBOL_STAR,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My third chart scatter'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_SCATTER, 'seriesName'=> ['2015', '2016'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'showSymbol' => false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>367
                ]
            ],
            [
                'title'=>Yii::t('app', 'My quarter chart scatter'),
                'series'=> [
                    ['October' => 12, 'November' => 15, 'December' => 13],
                    ['October' => 20, 'November' => 18, 'December' => 15],
                ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_SCATTER, 'seriesName'=> ['2015', '2016'],
                    'showSeriesName'=> false, 'showValue'=>true, 'showPercente'=>false,
                    'showSymbol' => true, 'typeSymbol' => PowerPoint::SYMBOL_DIAMOND,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>367
                ]
            ]
        ],
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Chart pie'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ]
        ],
        'charts' => [
            [
                'title'=>Yii::t('app', 'My first chart pie'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7],
                'options'=>['type'=>PowerPoint::CHART_TYPE_PIE,'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117]
            ],
            [
                'title'=>Yii::t('app', 'My second chart pie'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17],
                'options'=>['type'=>PowerPoint::CHART_TYPE_PIE,'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>117]
            ],
            [
                'title'=>Yii::t('app', 'My third chart pie'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7],
                'options'=>['type'=>PowerPoint::CHART_TYPE_PIE,'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>367]
            ],
            [
                'title'=>Yii::t('app', 'My quarter chart pie'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>['type'=>PowerPoint::CHART_TYPE_PIE,'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>367]
            ]
        ],
    ],
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Chart pie 3D'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ]
        ],
        'charts' => [
            [
                'title'=>Yii::t('app', 'My first chart pie 3D'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7],
                'options'=>['type'=>PowerPoint::CHART_TYPE_PIE_3D,'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117]
            ],
            [
                'title'=>Yii::t('app', 'My second chart pie 3D'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17],
                'options'=>['type'=>PowerPoint::CHART_TYPE_PIE_3D,'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>117]
            ],
            [
                'title'=>Yii::t('app', 'My third chart pie 3D'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7],
                'options'=>['type'=>PowerPoint::CHART_TYPE_PIE_3D,'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>367]
            ],
            [
                'title'=>Yii::t('app', 'My quarter chart pie 3D'),
                'series'=>['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>['type'=>PowerPoint::CHART_TYPE_PIE_3D,'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>367]
            ]
        ],
    ],
];

//generate ppt
Yii::app()->ppt->generate($options, $slides);