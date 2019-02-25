//
//  ViewController.m
//  ScrollMap
//
//  Created by kevin.chen on 2019/1/30.
//  Copyright © 2019 kevin.chen. All rights reserved.
//

#import "ViewController.h"
#import <MAMapKit/MAMapKit.h>
#import <AMapFoundationKit/AMapFoundationKit.h>
#import "CWMapView.h"
#import <AMapSearchKit/AMapSearchKit.h>
#import "xlsxwriter.h"

@interface ViewController ()<AMapSearchDelegate>

@property(nonatomic,strong)AMapSearchAPI *search;
// 已加载到的行数
@property (nonatomic, assign) int rowNum;

/**
 “小计”所在的单元格
 */
@property (nonatomic, retain) NSMutableArray *sumArray;

@end

static lxw_workbook  *workbook;
static lxw_worksheet *worksheet;

static lxw_format *titleformat;// 各表格标题栏的格式
static lxw_format *leftcontentformat;// 最左侧一列内容的样式
static lxw_format *contentformat;// 内容的样式
static lxw_format *rightcontentformat;// 最右侧一列内容的样式
static lxw_format *leftsumformat;// 最左侧一列小计的样式
static lxw_format *sumformat;// 小计的样式
static lxw_format *rightsumformat;// 最右侧一列小计的样式

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    // Do any additional setup after loading the view, typically from a nib.
    [AMapServices sharedServices].apiKey = @"cc529a292c1f593e24bbccaa75546134";
    [AMapServices sharedServices].enableHTTPS = YES;
    [self searchPOIData];

}
- (void)searchPOIData{
    self.search = [[AMapSearchAPI alloc] init];
    self.search.delegate = self;
    
    AMapPOIAroundSearchRequest *request = [[AMapPOIAroundSearchRequest alloc] init];
    
    request.location            = [AMapGeoPoint locationWithLatitude:31.221292 longitude:121.469376];
    request.keywords            = @"停车场";
    request.radius = 500;
    request.city = @"上海";
    /* 按照距离排序. */
    request.sortrule            = 0;
    request.requireExtension    = YES;
    
    [self.search AMapPOIAroundSearch:request];
}
/* POI 搜索回调. */
- (void)onPOISearchDone:(AMapPOISearchBaseRequest *)request response:(AMapPOISearchResponse *)response
{
    if (response.pois.count == 0)
    {
        return;
    }
    
    //解析response获取POI信息，具体解析见 Demo
    
    NSMutableArray *trafficArray = [NSMutableArray array];
    for (AMapPOI *poi in response.pois) {
//        NSLog(@"%@",poi.uid);
        
        NSDictionary *dic = @{
                              @"poiID": poi.uid,
                              @"name": poi.name,
                              @"location": [NSString stringWithFormat:@"%f,%f",poi.location.latitude,poi.location.longitude],
                              @"address":poi.address,
                              @"distance":[NSString stringWithFormat:@"%ld",(long)poi.distance],
                              @"parkingType":poi.parkingType,
                              @"image":[NSString stringWithFormat:@"%lu",(unsigned long)poi.images.count],
                              @"tel":poi.tel
                              };
        [trafficArray addObject:dic];
    }
    NSDictionary *dataDic = @{
                              @"traffic": trafficArray
                              };
    [self createXlsxFileWith:dataDic];
}
-(void)createXlsxFileWith:(NSDictionary *)dataDic{
    self.rowNum = 0;
    
    // 文件保存的路径
    NSString *documentPath = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory,NSUserDomainMask, YES) objectAtIndex:0];
    NSString *filename = [documentPath stringByAppendingPathComponent:@"c_demo.xlsx"];
    NSLog(@"filename_path:%@",filename);
    workbook  = workbook_new([filename UTF8String]);// 创建新xlsx文件，路径需要转成c字符串
    worksheet = workbook_add_worksheet(workbook, NULL);// 创建sheet
    [self setupFormat];
    
    //    [self createFormHeaderWithUserinfo:dataDic[@"userinfo"]];
    [self createTrafficForm:dataDic[@"traffic"]];
    //    [self createMealsForm:dataDic[@"meals"]];
    //    [self createTravelForm:dataDic[@"travel"]];
    //    [self creatOtherAndSumForm:dataDic[@"other"]];
    
    workbook_close(workbook);
}
// 单元格样式
-(void)setupFormat{
    titleformat = workbook_add_format(workbook);
    format_set_bold(titleformat);
    format_set_font_size(titleformat, 10);
    format_set_align(titleformat, LXW_ALIGN_CENTER);
    format_set_align(titleformat, LXW_ALIGN_VERTICAL_CENTER);//垂直居中
    format_set_border(titleformat, LXW_BORDER_NONE);// 边框（四周）：中宽边框
    
    leftcontentformat = workbook_add_format(workbook);
    format_set_font_size(leftcontentformat, 10);
    format_set_left(leftcontentformat, LXW_BORDER_NONE);// 左边框：中宽边框
    format_set_bottom(leftcontentformat, LXW_BORDER_NONE);// 下边框：双线边框
    
    contentformat = workbook_add_format(workbook);
    format_set_font_size(contentformat, 10);
    format_set_left(contentformat, LXW_BORDER_NONE);// 左边框：双线边框
    format_set_bottom(contentformat, LXW_BORDER_NONE);// 下边框：双线边框
    format_set_right(contentformat, LXW_BORDER_NONE);// 右边框：双线边框
    
    rightcontentformat = workbook_add_format(workbook);
    format_set_font_size(rightcontentformat, 10);
    format_set_bottom(rightcontentformat, LXW_BORDER_NONE);// 下边框：双线边框
    format_set_right(rightcontentformat, LXW_BORDER_NONE);// 右边框：中宽边框
//    format_set_num_format(rightcontentformat, "#,##0.00");
    
    leftsumformat = workbook_add_format(workbook);
    format_set_font_size(leftsumformat, 10);
    format_set_left(leftsumformat, LXW_BORDER_NONE);// 左边框：中宽边框
    format_set_bottom(leftsumformat, LXW_BORDER_NONE);// 下边框：中宽边框
    
    sumformat = workbook_add_format(workbook);
    format_set_font_size(sumformat, 10);
    format_set_align(sumformat, LXW_ALIGN_RIGHT);// 右对齐
    format_set_left(sumformat, LXW_BORDER_NONE);// 左边框：双线边框
    format_set_bottom(sumformat, LXW_BORDER_NONE);// 下边框：中宽边框
    format_set_right(sumformat, LXW_BORDER_NONE);// 右边框：双线边框
    
    rightsumformat = workbook_add_format(workbook);
    format_set_font_size(rightsumformat, 10);
    format_set_align(rightsumformat, LXW_ALIGN_RIGHT);// 右对齐
    format_set_bottom(rightsumformat, LXW_BORDER_NONE);// 下边框：中宽边框
    format_set_right(rightsumformat, LXW_BORDER_NONE);// 右边框：中宽边框
    format_set_num_format(rightsumformat, "#,##0.00");
}
// 整个文档的表头
-(void)createFormHeaderWithUserinfo:(NSDictionary *)userinfoDic{
    // 这个表格header标题格式
    lxw_format *headerFormat = workbook_add_format(workbook);
    format_set_font_size(headerFormat, 12);
    format_set_bold(headerFormat);
    format_set_align(headerFormat, LXW_ALIGN_CENTER);//水平居中
    format_set_align(headerFormat, LXW_ALIGN_VERTICAL_CENTER);//垂直居中
    
    
    // 姓名、报销日期格式
    lxw_format *nameFormat = workbook_add_format(workbook);
    format_set_font_size(nameFormat, 10);
    format_set_bold(nameFormat);
    
    // 设置列宽
    worksheet_set_column(worksheet, 3, 6, 60, NULL);// B、C两列宽度
    worksheet_set_column(worksheet, 3, 3, 25, NULL);// D列宽度
    
//    worksheet_write_string(worksheet, self.rowNum, 2, "月报销申请表", headerFormat);
//    worksheet_write_string(worksheet, ++self.rowNum, 0, "", NULL);//空白行
//    NSString *username = [NSString stringWithFormat:@"申报人：%@", userinfoDic[@"username"]];
//    const char *username_c = [username cStringUsingEncoding:NSUTF8StringEncoding];
//    worksheet_write_string(worksheet, ++self.rowNum, 1, username_c, nameFormat);
//    const char *dateRange_c = [userinfoDic[@"dateRange"] cStringUsingEncoding:NSUTF8StringEncoding];
//    worksheet_write_string(worksheet, self.rowNum, 3, dateRange_c, nameFormat);
}
// 市内交通费表格
-(void)createTrafficForm:(NSArray *)dataArray{
    [self setupFormContent:dataArray titleString:@"淮海路"];
}
//// 市内餐费表格
//-(void)createMealsForm:(NSArray *)dataArray{
//    [self setupFormContent:dataArray titleString:@"市内餐费"];
//}
//// 差旅费表格
//-(void)createTravelForm:(NSArray *)dataArray{
//    [self setupFormContent:dataArray titleString:@"差旅费"];
//}
//// 其他费用、合计
//-(void)creatOtherAndSumForm:(NSArray *)dataArray{
//    [self setupFormContent:dataArray titleString:@"其他费用"];
//}

-(void)setupFormContent:(NSArray *)dataArray titleString:(NSString *)titleString{
    worksheet_merge_range(worksheet, ++self.rowNum, 1, self.rowNum, 3, [titleString cStringUsingEncoding:NSUTF8StringEncoding], titleformat);
    if (![titleString isEqualToString:@"其他费用"]) {
        worksheet_write_string(worksheet, ++self.rowNum, 1, "POIID", titleformat);
        worksheet_write_string(worksheet, self.rowNum, 2, "名称", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 3, "兴趣点类型", titleformat);
//
//        worksheet_write_string(worksheet, self.rowNum, 4, "类型编码", titleformat);
        worksheet_write_string(worksheet, self.rowNum, 3, "经纬度", titleformat);
        worksheet_write_string(worksheet, self.rowNum, 4, "地址", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 7, "电话", titleformat);
        worksheet_write_string(worksheet, self.rowNum, 5, "距中心点的距离", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 9, "停车场类型", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 10, "商铺id", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 11, "邮编", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 12, "网址", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 13, "电子邮件", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 14, "省", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 15, "省编码", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 16, "城市名称", titleformat);
//
//        worksheet_write_string(worksheet, self.rowNum, 17, "城市编码", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 18, "区域名称", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 19, "区域编码", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 20, "地理格ID", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 21, "入口经纬度", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 22, "出口经纬度", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 23, "方向", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 24, "是否有室内地图", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 25, "所在商圈", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 26, "室内信息", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 27, "子POI列表", titleformat);
        worksheet_write_string(worksheet, self.rowNum, 6, "图片张数", titleformat);
//        worksheet_write_string(worksheet, self.rowNum, 29, "扩展信息", titleformat);
        worksheet_write_string(worksheet, self.rowNum, 7, "停车场类型", titleformat);
        
        worksheet_write_string(worksheet, self.rowNum, 8, "电话", titleformat);
        
    }
    
    int startRow = self.rowNum;
//    @"poiID": poi.uid,
//    @"name": poi.name,
//    @"location": [NSString stringWithFormat:@"%f,%f",poi.location.latitude,poi.location.longitude],
//    @"address":poi.address,
//    @"distance":[NSString stringWithFormat:@"%ld",(long)poi.distance],
//    @"image":[NSString stringWithFormat:@"%lu",(unsigned long)poi.images.count]
    for (int i = 0; i < dataArray.count; i++) {
        NSDictionary *dic = dataArray[i];
        worksheet_write_string(worksheet, ++self.rowNum, 1, [dic[@"poiID"] cStringUsingEncoding:NSUTF8StringEncoding], leftcontentformat);
        worksheet_write_string(worksheet, self.rowNum, 2,  [dic[@"name"] cStringUsingEncoding:NSUTF8StringEncoding], contentformat);
        worksheet_write_string(worksheet, self.rowNum, 3,  [dic[@"location"] cStringUsingEncoding:NSUTF8StringEncoding], contentformat);
//        worksheet_write_number(worksheet, self.rowNum, 3, [dic[@"location"] doubleValue], rightcontentformat);
        worksheet_write_string(worksheet, self.rowNum, 4,  [dic[@"address"] cStringUsingEncoding:NSUTF8StringEncoding], rightcontentformat);
//        worksheet_write_number(worksheet, self.rowNum, 4, dic[@"address"], rightcontentformat);
        worksheet_write_number(worksheet, self.rowNum, 5, [dic[@"distance"] doubleValue], rightcontentformat);
        worksheet_write_number(worksheet, self.rowNum, 6, [dic[@"image"] doubleValue], rightcontentformat);
        
        worksheet_write_string(worksheet, self.rowNum, 7,  [dic[@"parkingType"] cStringUsingEncoding:NSUTF8StringEncoding], rightcontentformat);
        
        worksheet_write_string(worksheet, self.rowNum, 8,  [dic[@"tel"] cStringUsingEncoding:NSUTF8StringEncoding], rightcontentformat);
    }
    // 空行
    //    worksheet_write_string(worksheet, ++self.rowNum, 1, "", leftcontentformat);
    //    worksheet_write_string(worksheet, self.rowNum, 2, "", contentformat);
    //    worksheet_write_number(worksheet, self.rowNum, 3, 0, rightcontentformat);
    //
    //    int endRow = self.rowNum;
    //    NSString *sumFormula = [NSString stringWithFormat:@"=SUM(D%d:D%d)", startRow+1, endRow+1];
    //    worksheet_write_string(worksheet, ++self.rowNum, 1, "", leftsumformat);
    //    worksheet_write_string(worksheet, self.rowNum, 2, "小计：", sumformat);
    //    worksheet_write_formula(worksheet, self.rowNum, 3, [sumFormula cStringUsingEncoding:NSUTF8StringEncoding], rightsumformat);
    //
    //    [self.sumArray addObject:@(self.rowNum+1)];// 记录小计金额单元格位置
    
    if ([titleString isEqualToString:@"其他费用"]) {
        //        [self sumTotalMoney];
    } else {
        worksheet_write_string(worksheet, ++self.rowNum, 0, "", NULL);// 空行
        worksheet_write_string(worksheet, ++self.rowNum, 0, "", NULL);// 空行
    }
}
-(void)sumTotalMoney{
    lxw_format *borderformat_alignleft = workbook_add_format(workbook);
    format_set_font_size(borderformat_alignleft, 10);
    format_set_border(borderformat_alignleft, LXW_BORDER_MEDIUM);//  边框（四周）：中宽边框
    worksheet_merge_range(worksheet, ++self.rowNum, 1, self.rowNum, 3, "备注：如有特殊说明请在此栏填写", borderformat_alignleft);
    
    lxw_format *borderformat_alignright = workbook_add_format(workbook);
    format_set_font_size(borderformat_alignright, 10);
    format_set_bold(borderformat_alignright);
    format_set_border(borderformat_alignright, LXW_BORDER_MEDIUM);//  边框（四周）：中宽边框
    format_set_align(borderformat_alignright, LXW_ALIGN_RIGHT);
    worksheet_write_string(worksheet, ++self.rowNum, 1, "", borderformat_alignleft);
    worksheet_write_string(worksheet, self.rowNum, 2, "合计", borderformat_alignright);
    
    lxw_format *totalmoneyformat = workbook_add_format(workbook);
    format_set_font_size(totalmoneyformat, 10);
    format_set_bold(totalmoneyformat);
    format_set_border(totalmoneyformat, LXW_BORDER_MEDIUM);//  边框（四周）：中宽边框
    format_set_align(totalmoneyformat, LXW_ALIGN_RIGHT);
    format_set_num_format(totalmoneyformat, "￥#,##0.00");
    NSString *sumStr = @"=D";
    for (int i = 0; i < self.sumArray.count; i++) {
        if (i < self.sumArray.count-1) {
            sumStr = [NSString stringWithFormat:@"%@%@+D", sumStr, self.sumArray[i]];
        } else {
            sumStr = [NSString stringWithFormat:@"%@%@", sumStr, self.sumArray[i]];
        }
    }
    NSLog(@"sumarray:%@", self.sumArray);
    worksheet_write_formula_num(worksheet, self.rowNum, 3, [sumStr cStringUsingEncoding:NSUTF8StringEncoding], totalmoneyformat, 0);
}
-(NSMutableArray *)sumArray{
    if (!_sumArray) {
        _sumArray = [NSMutableArray array];
    }
    return _sumArray;
}
- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}


@end
