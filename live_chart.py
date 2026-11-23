"""Native Tk Canvas live chart: no browser, JavaScript, or extra dependency."""
import math
from bisect import bisect_left
from datetime import datetime


def clean_points(samples, now, window, factor=1.0):
    points = []
    for row in samples:
        try:
            ts, value = float(row[0]), float(row[1]) * factor
            if math.isfinite(ts) and math.isfinite(value) and value > 0 and now-window <= ts <= now+1:
                points.append((ts, value))
        except (TypeError, ValueError, IndexError, OverflowError):
            continue
    points.sort()
    return points


def reduce_points(points, limit=600):
    """Preserve first/last and min/max in chronological buckets."""
    if len(points) <= limit:
        return list(points)
    result = [points[0]]
    bucket = max(1, math.ceil((len(points)-2) / ((limit-2)//2)))
    for start in range(1, len(points)-1, bucket):
        segment = points[start:min(start+bucket, len(points)-1)]
        lo = min(range(len(segment)), key=lambda i: segment[i][1])
        hi = max(range(len(segment)), key=lambda i: segment[i][1])
        result.extend(segment[i] for i in sorted(set((lo, hi))))
    result.append(points[-1])
    return result


def render_chart(canvas, points, now, window, theme, unit, privacy=False,
                 average=None, transactions=None, viewport=None, font_size=12):
    font_size = max(10, min(18, int(font_size)))
    transactions = list(transactions or [])
    canvas.delete('all')
    width, height = canvas.winfo_width(), canvas.winfo_height()
    if width < 220 or height < 150:
        return None
    muted = theme.get('muted', '#8d98a8')
    text = theme.get('text', '#edf2f7')
    border = theme.get('border', '#273244')
    if privacy:
        canvas.create_text(width/2, height/2, text='وضع الخصوصية — الرسم مخفي', fill=muted,
                           font=('Segoe UI', font_size))
        return None
    left, right, top, bottom = 55+font_size*4, width-20, 40, height-45
    oldest, newest = now-window, now
    values = [p[1] for p in points]
    if not values and not transactions:
        canvas.create_text(width/2, height/2, text='بانتظار أول تسعير حديث…', fill=muted,
                           font=('Segoe UI', font_size))
        return None
    # Fit available history when a day/week/month has not accumulated yet.
    if viewport:
        oldest,newest = viewport
        span = max(10,newest-oldest)
    else:
        newest = max(now, points[-1][0] if points else now)
        first = min([p[0] for p in points]+[r['ts'] for r in transactions])
        span = min(window, max(60.0, newest-first+5.0))
        oldest = newest-span
    transactions = [r for r in transactions if oldest <= r['ts'] <= newest]
    scale_values = values+[r['value'] for r in transactions]
    if average is not None and math.isfinite(average) and average>0:
        scale_values.append(average)
    else:
        average=None
    low, high = min(scale_values), max(scale_values)
    pad = max((high-low)*.15, abs(scale_values[-1])*.00003, .003)
    low, high = low-pad, high+pad
    def x(ts):
        return left+(ts-oldest)/span*(right-left)
    def y(value):
        return top+(high-value)/(high-low)*(bottom-top)
    for index in range(5):
        value = low+(high-low)*index/4
        yy = y(value)
        canvas.create_line(left, yy, right, yy, fill=border)
        canvas.create_text(left-9, yy, text=f'{value:,.3f}', anchor='e', fill=muted,
                           font=('Segoe UI', font_size-1))
    ticks = max(2, min(5, int((right-left)/max(170,font_size*14))+1))
    for index in range(ticks):
        ts = oldest+span*index/(ticks-1)
        anchor = 'w' if index==0 else 'e' if index==ticks-1 else 'center'
        fmt = '%m-%d' if span>7*86400 else '%m-%d %H:%M' if span>86400 else '%H:%M:%S' if span<=900 else '%H:%M'
        canvas.create_line(x(ts),top,x(ts),bottom,fill=border)
        canvas.create_text(x(ts), bottom+18, text=datetime.fromtimestamp(ts).strftime(fmt),
                           anchor=anchor, fill=muted, font=('Segoe UI', font_size-1))
    canvas.create_text(left, 14, text=unit, anchor='w', fill=text, font=('Segoe UI', font_size, 'bold'))
    color = theme.get('profit', '#2ecc71') if not values or values[-1] >= values[0] else theme.get('loss', '#ff5c5c')
    # Deliberately no smooth interpolation or invented intermediate samples.
    reduced = reduce_points(points)
    groups = [[]]
    for point in reduced:
        if groups[-1] and point[0]-groups[-1][-1][0] > max(20, window/200):
            groups.append([])
        groups[-1].append(point)
    for group in groups:
        if len(group)>1:
            coords = [coord for ts,v in group for coord in (x(ts),y(v))]
            canvas.create_line(*coords, fill=color, width=2.5, smooth=False)
        elif group:
            xx,yy=x(group[0][0]),y(group[0][1])
            canvas.create_oval(xx-2,yy-2,xx+2,yy+2,fill=color,outline='')
    if values:
        last_x,last_y=x(points[-1][0]),y(values[-1])
        canvas.create_line(left,last_y,right,last_y,fill=color,dash=(4,4))
        canvas.create_oval(last_x-4,last_y-4,last_x+4,last_y+4,fill=color,outline='')
        canvas.create_text(right,16,text=f'{values[-1]:,.3f}',anchor='e',fill=color,
                           font=('Segoe UI',font_size,'bold'))
    if average is not None:
        yy=y(average)
        canvas.create_line(left,yy,right,yy,fill='#d6a33a',dash=(8,4),width=2,tags='average')
        canvas.create_text(left+5,yy-10,text=f'متوسط التكلفة {average:,.3f}',anchor='w',
                           fill=theme.get('text','#edf2f7'),font=('Segoe UI',font_size-1,'bold'),tags='average')
    markers=[]
    for row in transactions:
        xx,yy=x(row['ts']),y(row['value'])
        fill=theme.get('profit','#2ecc71') if row['type']=='buy' else theme.get('loss','#ff5c5c')
        canvas.create_oval(xx-7,yy-7,xx+7,yy+7,fill=fill,outline=text,width=1,tags='transaction')
        canvas.create_text(xx,yy,text='B' if row['type']=='buy' else 'S',fill=theme.get('card','#12161e'),
                           font=('Segoe UI',9,'bold'),tags='transaction')
        markers.append((xx,yy,row))
    return {'points':points,'xs':[x(p[0]) for p in points], 'x':x,'y':y,
            'bounds':(left,top,right,bottom),'color':color,'unit':unit,'span':span,
            'oldest':oldest,'newest':newest,'markers':markers,'font_size':font_size}


def hover_chart(canvas, layout, event, theme):
    canvas.delete('hover')
    if not layout:
        return
    left,top,right,bottom=layout['bounds']
    if not left <= event.x <= right or not top <= event.y <= bottom:
        return
    markers=[r for r in layout.get('markers',[]) if abs(r[0]-event.x)<=12 and abs(r[1]-event.y)<=12]
    if markers:
        row=min(markers,key=lambda r:(r[0]-event.x)**2+(r[1]-event.y)**2)[2]
        message=row['label']
    else:
        xs=layout['xs']
        index=bisect_left(xs,event.x)
        candidates=[i for i in (index-1,index) if 0 <= i < len(xs)]
        if not candidates:
            return
        nearest=min(candidates,key=lambda i:abs(xs[i]-event.x))
        ts,value=layout['points'][nearest]
        xx,yy=layout['x'](ts),layout['y'](value)
        canvas.create_line(xx,top,xx,bottom,fill=theme.get('muted','#8d98a8'),dash=(3,3),tags='hover')
        canvas.create_oval(xx-5,yy-5,xx+5,yy+5,outline=layout['color'],width=2,tags='hover')
        message=f'{datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")}  |  {value:,.3f}'
    muted=theme.get('muted','#8d98a8')
    tip=canvas.create_text((left+right)/2,top+12,text=message,fill=theme.get('text','#edf2f7'),
                           font=('Segoe UI',layout.get('font_size',12),'bold'),width=right-left-12,tags='hover')
    box=canvas.bbox(tip)
    if box:
        bg=canvas.create_rectangle(box[0]-6,box[1]-3,box[2]+6,box[3]+3,
                                   fill=theme.get('card','#12161e'),outline=muted,tags='hover')
        canvas.tag_lower(bg,tip)
