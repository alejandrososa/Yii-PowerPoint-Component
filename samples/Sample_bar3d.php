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
                'text'=>Yii::t('app', 'Delivery'),
                'options'=>[
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', 'Abril/2016'),
                'options'=>[
                    'height'=>38, 'width'=>132, 'ox'=>816, 'oy'=>52,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', 'State Budget 2016'),
                'options'=>[
                    'height'=>38, 'width'=>403, 'ox'=>326, 'oy'=>73,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
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
    ]
];

//generate ppt
Yii::app()->ppt->generate($options, $slides);