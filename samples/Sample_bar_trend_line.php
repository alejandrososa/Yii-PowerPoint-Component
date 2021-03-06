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
                'title'=>Yii::t('app', 'My first chart bar'),
                'series'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 7, 'Tuesday' => 9, 'Wednesday' => 19, 'Thursday' => 16, 'Friday' => 8, 'Saturday' => 11],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'seriesTrends'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 7, 'Tuesday' => 9, 'Wednesday' => 19, 'Thursday' => 16, 'Friday' => 8, 'Saturday' => 11],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_TRENDLINE,
                    'trendsName'=> ['less', 'equal', 'greater'], 'showTrendsName'=> false,
                    'seriesName'=> ['less', 'equal', 'greater'], 'showSeriesName'=> false,
                    'showValue'=>true, 'showSymbol' => false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117
                ]
            ],
            [
                'title'=>Yii::t('app', 'My second chart bar'),
                'series'=> ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'seriesTrends'=> ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_TRENDLINE,
                    'trendsName'=> ['week'], 'showTrendsName'=> false,
                    'seriesName'=> ['week'], 'showSeriesName'=> false,
                    'titleItalic' => true, 'legendItalic' => true,
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
                'seriesTrends'=>
                    [
                        ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9],
                        ['Monday' => 12, 'Tuesday' => 20, 'Wednesday' => 13, 'Thursday' => 7, 'Friday' => 4, 'Saturday' => 9]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_TRENDLINE,
                    'trendsName'=> ['2009', '2010'], 'showTrendsName'=> true,
                    'seriesName'=> ['2009', '2010'], 'showSeriesName'=> true,
                    'titleItalic' => true, 'legendItalic' => true,
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
                'seriesTrends'=>
                    [
                        ['October' => 12, 'November' => 15, 'December' => 13],
                        ['October' => 11, 'November' => 20, 'December' => 14]
                    ],
                'options'=>[
                    'type'=>PowerPoint::CHART_TYPE_BAR_TRENDLINE,
                    'trendsName'=> ['2014', '2015'], 'showTrendsName'=> true,
                    'seriesName'=> ['2014', '2015'], 'showSeriesName'=> true,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>367
                ]
            ]
        ],
    ]
];

//generate ppt
Yii::app()->ppt->generate($options, $slides);