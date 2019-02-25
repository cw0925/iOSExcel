//
//  CWMapView.m
//  ScrollMap
//
//  Created by kevin.chen on 2019/1/30.
//  Copyright © 2019 kevin.chen. All rights reserved.
//

#import "CWMapView.h"
#import <MAMapKit/MAMapKit.h>
#import <AMapFoundationKit/AMapFoundationKit.h>

@implementation CWMapView

- (instancetype)initWithFrame:(CGRect)frame{
    self = [super initWithFrame:frame];
    if (self) {
        ///初始化地图
        MAMapView *_mapView = [[MAMapView alloc] initWithFrame:frame];
        
        ///把地图添加至view
        [self addSubview:_mapView];
    }
    return self;
}

@end
