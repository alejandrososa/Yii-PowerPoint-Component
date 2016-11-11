<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 20/09/2016
 * Hora: 15:14
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;
use AlejandroSosa\YiiPowerPoint\Common\Helper;

/**
 * Class Charts
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Charts extends AbstractObject
{
    /**
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function create(Slide $slide, $options = array())
    {
        //check if options is only one or multiple
        if (Helper::isArrayMultidimensional($options)) {
            foreach ($options as $item) {
                self::createCustomChart($slide, $item);
            }
        } else {
            self::createCustomChart($slide, $options);
        }
    }

    /**
     * Creates different types StatGraphs
     * @param Slide $slide
     * @param array $params
     */
    private function createCustomChart(Slide $slide, $params = array())
    {
        $options    = Helper::hasArrayProperty('options', $params) ? $params['options'] : array();
        $title      = Helper::hasArrayProperty('title', $params) ? $params['title'] : '';
        $data       = Helper::hasArrayProperty('series', $params) ? $params['series'] : array();
        $type       = Helper::hasArrayProperty('type', $options) ? $options['type'] : '';

        switch ($type) {
            case self::CHART_TYPE_AREA:
                //TODO library has bug for chart area
                //self::makeArea($slide, $title, $data, $options); break;
            case self::CHART_TYPE_LINE:
                break;
            case self::CHART_TYPE_BAR:
                self::makeBar($slide, self::CHART_TYPE_BAR, $title, $data, $options);
                break;
            case self::CHART_TYPE_BAR_HORIZONTAL:
                self::makeBar($slide, self::CHART_TYPE_BAR_HORIZONTAL, $title, $data, $options);
                break;
            case self::CHART_TYPE_BAR_STACKED:
                self::makeBar($slide, self::CHART_TYPE_BAR_STACKED, $title, $data, $options);
                break;
            case self::CHART_TYPE_BAR_PERCENT_STACKED:
                self::makeBar($slide, self::CHART_TYPE_BAR_PERCENT_STACKED, $title, $data, $options);
                break;
            case self::CHART_TYPE_BAR_3D:
                self::makeBar($slide, self::CHART_TYPE_BAR_3D, $title, $data, $options);
                break;
            case self::CHART_TYPE_PIE:
                self::makePie($slide, self::CHART_TYPE_PIE, $title, $data, $options);
                break;
            case self::CHART_TYPE_PIE_3D:
                self::makePie($slide, self::CHART_TYPE_PIE_3D, $title, $data, $options);
                break;
            case self::CHART_TYPE_SCATTER:
                self::makeScatter($slide, $title, $data, $options);
                break;
        }
    }


    /**
     * Make chart Bar
     * @param Slide $slide
     * @param $type_bar
     * @param string $title
     * @param array $series_data
     * @param array $options
     */
    private function makeBar(Slide $slide, $type_bar, $title = '', $series_data = array(), $options = array())
    {
        $series_name      = Helper::hasArrayProperty('seriesName', $options) ? $options['seriesName'] : '';
        $show_percent     = Helper::hasArrayProperty('showPercente', $options) ? $options['showPercente'] : true;
        $show_value       = Helper::hasArrayProperty('showValue', $options) ? $options['showValue'] : true;
        $show_series_name = Helper::hasArrayProperty('showSeriesName', $options) ? $options['showSeriesName'] : true;

        if ($slide instanceof Slide) {
            // Create a bar chart (that should be inserted in a shape)
            $barChart = $type_bar == self::CHART_TYPE_BAR_3D ? new Bar3D() : new Bar();

            //set value from serie to percentage
            if($type_bar == self::CHART_TYPE_BAR_PERCENT_STACKED){
                $series_data = self::convertSeriePercentage($series_data);
                $options_serie['format']  = self::TEXT_FORMAT_PERCENTE;
            }

            //set angle 3d
            if($type_bar == self::CHART_TYPE_BAR_3D){
                $options_serie['angleAxes']  = true;
            }

            //check if series is only one or multiple
            if (Helper::isArrayMultidimensional($series_data)) {
                foreach ($series_data as $index => $item) {
                    //create new serie
                    $options_serie['showPercente']  = $show_percent;
                    $options_serie['showName']      = $show_series_name;
                    $options_serie['showValue']     = $show_value;
                    $options_serie['title']         = is_array($series_name) && !empty($series_name[$index])
                        ? $series_name[$index] : '';
                    $serie = self::createNewSerie($item, $options_serie);

                    //add serie to barChart
                    $barChart->addSeries($serie);
                }
            } else {
                //create new serie
                $options_serie['showPercente']  = $show_percent;
                $options_serie['showName']      = $show_series_name;
                $options_serie['showValue']     = $show_value;
                $options_serie['title']         = is_array($series_name) ? $series_name[0] : $series_name;
                $serie                          = self::createNewSerie($series_data, $options_serie);

                //add serie to barChart
                $barChart->addSeries($serie);
            }

            //set type of bar
            self::setTypeBar($barChart, $type_bar);

            // Create a shape (chart)
            self::createNewShapeChart($slide, $barChart, $title, $options);
        }
    }

    /**
     * Make chart Pie
     * @param Slide $slide
     * @param string $title
     * @param array $series_data
     * @param array $options
     */
    private function makePie(Slide $slide, $type_bar, $title = '', $series_data = array(), $options = array())
    {
        $series_name      = Helper::hasArrayProperty('seriesName', $options) ? $options['seriesName'] : '';
        $show_percent     = Helper::hasArrayProperty('showPercente', $options) ? $options['showPercente'] : true;
        $show_value       = Helper::hasArrayProperty('showValue', $options) ? $options['showValue'] : true;
        $show_series_name = Helper::hasArrayProperty('showSeriesName', $options) ? $options['showSeriesName'] : true;

        if($slide instanceof Slide) {
            //create a pie chart (that should be inserted in a shape)
            if($type_bar == self::CHART_TYPE_PIE){
                $pieChart = new Pie();
            }else if($type_bar == self::CHART_TYPE_PIE_3D){
                $pieChart = new Pie3D();
            }
            $pieChart->setExplosion(20);

            //create new serie
            $options_serie = array();
            $options_serie['showPercente']  = $show_percent;
            $options_serie['showName']      = $show_series_name;
            $options_serie['showValue']     = $show_value;
            $options_serie['title']         = is_array($series_name) ? $series_name[0] : $series_name;
            $serie                          = self::createNewSerie($series_data, $options_serie);

            //add serie to chart
            $pieChart->addSeries($serie);

            //create a shape (chart)
            self::createNewShapeChart($slide, $pieChart, $title, $options);
        }
    }

    /**
     * Make chart Scatter
     * @param Slide $slide
     * @param string $title
     * @param array $series_data
     * @param array $options
     */
    private function makeScatter(Slide $slide, $title = '', $series_data = array(), $options = array())
    {
        $series_name      = Helper::hasArrayProperty('seriesName', $options) ? $options['seriesName'] : '';
        $show_percent     = Helper::hasArrayProperty('showPercente', $options) ? $options['showPercente'] : true;
        $show_value       = Helper::hasArrayProperty('showValue', $options) ? $options['showValue'] : true;
        $show_series_name = Helper::hasArrayProperty('showSeriesName', $options) ? $options['showSeriesName'] : true;
        $show_symbol      = Helper::hasArrayProperty('showSymbol', $options) ? $options['showSymbol'] : true;
        $type_symbol      = Helper::hasArrayProperty('typeSymbol', $options) ? $options['typeSymbol'] : '';

        if($slide instanceof Slide) {
            //create a pie chart (that should be inserted in a shape)
            $scatterChart = new Scatter();
            $options_serie['showSymbol']    = $show_symbol;
            $options_serie['typeSymbol']    = $type_symbol;

            //check if series is only one or multiple
            if (Helper::isArrayMultidimensional($series_data)) {
                foreach ($series_data as $index => $item) {
                    //create new serie
                    $options_serie = array();
                    $options_serie['showPercente']  = $show_percent;
                    $options_serie['showName']      = $show_series_name;
                    $options_serie['showValue']     = $show_value;
                    $options_serie['title']         = is_array($series_name) && !empty($series_name[$index]) ? $series_name[$index] : '';
                    $serie                          = self::createNewSerie($item, $options_serie);

                    //add serie to chart
                    $scatterChart->addSeries($serie);
                }
            } else {
                //create new serie
                $options_serie = array();
                $options_serie['showPercente']  = $show_percent;
                $options_serie['showName']      = $show_series_name;
                $options_serie['showValue']     = $show_value;
                $options_serie['title']         = is_array($series_name) ? $series_name[0] : $series_name;
                $serie                          = self::createNewSerie($series_data, $options_serie);

                //add serie to chart
                $scatterChart->addSeries($serie);
            }

            //create a shape (chart)
            self::createNewShapeChart($slide, $scatterChart, $title, $options);
        }
    }


    /**
     * Set color of shape chart
     * @param $color
     * @return Fill
     */
    private function getFill($color = self::COLOR_WHITE)
    {
        $oFill = new Fill();

        if($color == self::COLOR_NONE){
            $oFill->setFillType(Fill::FILL_NONE)->setStartColor(new Color($color));
        }else{
            $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($color));
        }

        return $oFill;
    }

    /**
     * Set shadow of shape chart
     * @param bool $visible
     * @param int $direction
     * @param int $distance
     * @return static
     */
    private function getShadow($visible = true, $direction = 45, $distance = 10)
    {
        $oShadow = new Shadow();
        return $oShadow->setVisible($visible)->setDirection($direction)->setDistance($distance);
    }

    /**
     * Create a series of array given
     * Example: ['Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14]
     * @param array $serie
     * @param array $options
     * @return Series
     */
    private function createNewSerie($serie = array(), $options = array())
    {
        $title      = Helper::hasArrayProperty('title', $options) ? $options['title'] : 'de';
        $color      = Helper::hasArrayProperty('color', $options) ? $options['color'] : '';
        $format     = Helper::hasArrayProperty('format', $options) ? $options['format'] : '';
        $show_name  = Helper::hasArrayProperty('showName', $options) ? $options['showName'] : true;
        $show_value = Helper::hasArrayProperty('showValue', $options) ? $options['showValue'] : true;
        $show_perc  = Helper::hasArrayProperty('showPercente', $options) ? $options['showPercente'] : true;
        $show_symbol= Helper::hasArrayProperty('showSymbol', $options) ? $options['showSymbol'] : false;
        $type_symbol= Helper::hasArrayProperty('typeSymbol', $options) ? $options['typeSymbol'] : '';

        $obj_serie = new Series($title, $serie);
        $obj_serie->setShowSeriesName($show_name);
        $obj_serie->setShowValue($show_value);
        $obj_serie->setShowPercentage($show_perc);
        $obj_serie->setLabelPosition(Series::LABEL_BESTFIT);
        $obj_serie->setDlblNumFormat($format);
        $obj_serie->getFont()->getColor()->setRGB(Style::getRGB(self::COLOR_PRIMARY_TEXT));

        //symbol
        if($show_symbol){
            $type_symbol = self::getSymbol($type_symbol);
            $obj_serie->getMarker()->setSymbol($type_symbol);
        }

        //color background
        if(!empty($color)){
            $obj_serie->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($color));
        }

        //TODO Highlight column with a specific color
        //$obj_serie->getFont()->getColor()->setRGB('00FF00');
        //$obj_serie->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
        return $obj_serie;
    }

    /**
     * Create shape chart
     * @param Slide $slide
     * @param $chart
     * @param $title
     * @param array $options
     */
    private function createNewShapeChart(Slide $slide, $chart, $title, $options = array())
    {
        $background     = Helper::hasArrayProperty('background', $options) ? $options['background'] : self::COLOR_NONE;
        $height         = Helper::hasArrayProperty('height', $options) ? $options['height'] : self::CHART_HEIGHT;
        $width          = Helper::hasArrayProperty('width', $options) ? $options['width'] : self::CHART_WIDTH;
        $offset_x       = Helper::hasArrayProperty('ox', $options) ? $options['ox'] : self::CHART_OFFSET_X;
        $offset_y       = Helper::hasArrayProperty('oy', $options) ? $options['oy'] : self::CHART_OFFSET_Y;
        $title_italic   = Helper::hasArrayProperty('titleItalic', $options) ? $options['titleItalic'] : false;
        $legend_italic  = Helper::hasArrayProperty('legendItalic', $options) ? $options['legendItalic'] : false;
        $r_angle_axes   = Helper::hasArrayProperty('angleAxes', $options) ? $options['angleAxes'] : false;
        $legend_offset_x= Helper::hasArrayProperty('legendOx', $options) ? $options['legendOx'] : self::CHART_OFFSET_X;
        $legend_offset_y= Helper::hasArrayProperty('legendOy', $options) ? $options['legendOy'] : self::CHART_OFFSET_Y;
        $legend_bg      = Helper::hasArrayProperty('legendBackground', $options) ? $options['legendBackground'] : self::COLOR_NONE;
        $legend_visible = Helper::hasArrayProperty('legendVisible', $options) ? $options['legendVisible'] : self::TRUE;


        //add background, border and shadow
        $oFill          = self::getFill($background);
        $oShadow        = self::getShadow(false);

        $shape = $slide->createChartShape();
        $shape->setName($title)
            ->setResizeProportional(false)
            ->setHeight($height)->setWidth($width)
            ->setOffsetX($offset_x)->setOffsetY($offset_y);
        $shape->setShadow($oShadow);
        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getTitle()->setText($title);
        $shape->getTitle()->getFont()->setItalic($title_italic);
        $shape->getPlotArea()->setType($chart);
        $shape->getView3D()->setRightAngleAxes($r_angle_axes);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getLegend()->getFont()->setItalic($legend_italic);
        $shape->getLegend()->setOffsetX($legend_offset_x);
        $shape->getLegend()->getOffsetY($legend_offset_y);
        $shape->getLegend()->setVisible($legend_visible);

        //bg legend
        if ($legend_bg == self::COLOR_NONE){
            $shape->getLegend()->getFill()->setFillType(Fill::FILL_NONE)->setStartColor(new Color($legend_bg));
        }else{
            $shape->getLegend()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($legend_bg));
        }
    }

    /**
     * Convert values from serie to percentage
     * @param array $serie
     * @return array
     */
    private function convertSeriePercentage($serie = array())
    {
        //check if multiple serie
        $check_serie = Helper::isArrayMultidimensional($serie);
        if ($check_serie) {
            //walk all serie
            $_series = $result = array();
            foreach ($serie as $item) {
                //set value to percentage
                $series_sum = array_sum($item);
                foreach ($item as $cat_name => $Value) {
                    $_series[$cat_name] = round($Value / $series_sum, 2);
                }
                $result[] = $_series;
            }
        }else{
            //set value to percentage
            $series_sum = array_sum($serie);
            foreach ($serie as $cat_name => $Value) {
                $result[$cat_name] = round($Value / $series_sum, 2);
            }
        }

        return $result;
    }

    /**
     * Set type of Bar
     * @param Bar $bar
     * @param $type
     */
    private function setTypeBar($bar, $type)
    {
        if($bar instanceof Bar || $bar instanceof Bar3D){
            switch ($type){
                case self::CHART_TYPE_BAR:
                    break;
                case self::CHART_TYPE_BAR_HORIZONTAL:
                    $bar->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);
                    break;
                case self::CHART_TYPE_BAR_STACKED:
                    $bar->setBarGrouping( Bar::GROUPING_STACKED );
                    break;
                case self::CHART_TYPE_BAR_PERCENT_STACKED:
                    $bar->setBarGrouping( Bar::GROUPING_PERCENTSTACKED );
                    $bar->setBarDirection( Bar3D::DIRECTION_HORIZONTAL );
                    break;
            }
        }
    }

    /**
     * Get symbol
     * @param $type
     * @return null|string
     */
    private function getSymbol($type)
    {
        $obj = null;
        switch ($type){
            case self::SYMBOL_CIRCLE:
                $obj = Marker::SYMBOL_CIRCLE;
                break;
            case self::SYMBOL_DASH:
                $obj = Marker::SYMBOL_DASH;
                break;
            case self::SYMBOL_DIAMOND:
                $obj = Marker::SYMBOL_DIAMOND;
                break;
            case self::SYMBOL_DOT:
                $obj = Marker::SYMBOL_DOT;
                break;
            case self::SYMBOL_NONE:
                $obj = Marker::SYMBOL_NONE;
                break;
            case self::SYMBOL_PLUS:
                $obj = Marker::SYMBOL_PLUS;
                break;
            case self::SYMBOL_SQUARE:
                $obj = Marker::SYMBOL_SQUARE;
                break;
            case self::SYMBOL_STAR:
                $obj = Marker::SYMBOL_STAR;
                break;
            case self::SYMBOL_TRIANGLE:
                $obj = Marker::SYMBOL_TRIANGLE;
                break;
            case self::SYMBOL_X:
                $obj = Marker::SYMBOL_X;
                break;
            default: $obj = Marker::SYMBOL_DOT;
        }
        return $obj;
    }
}