package com.inzent.yellowpage.migration ;

import org.springframework.beans.factory.annotation.Autowired ;
import org.springframework.stereotype.Component ;

import com.inzent.yellowpage.marshaller.PublishingFormat ;
import com.inzent.yellowpage.model.PublishLog ;
import com.inzent.yellowpage.publish.PublishingFormatMarshallerBean ;

@Component
public class MigrationBean extends Migration
{
  @Autowired
  protected PublishingFormatMarshallerBean publishFormatMarshallerBean ;

  @Override
  protected PublishLog unmarshal(PublishingFormat publishingFormat)
  {
    return publishFormatMarshallerBean.unmarshal(publishingFormat) ;
  }
}
