package com.inzent.yellowpage.publish ;

import java.io.ByteArrayInputStream ;
import java.io.ByteArrayOutputStream ;
import java.io.IOException ;
import java.io.StringWriter ;
import java.sql.Timestamp ;
import java.util.Objects ;

import org.apache.commons.lang3.StringUtils ;
import org.apache.commons.logging.Log ;
import org.apache.commons.logging.LogFactory ;
import org.apache.hc.client5.http.classic.methods.HttpDelete ;
import org.apache.hc.client5.http.classic.methods.HttpPost ;
import org.apache.hc.client5.http.classic.methods.HttpUriRequestBase ;
import org.apache.hc.client5.http.impl.classic.AbstractHttpClientResponseHandler ;
import org.apache.hc.client5.http.impl.classic.HttpClientBuilder ;
import org.apache.hc.client5.http.impl.io.BasicHttpClientConnectionManager ;
import org.apache.hc.core5.http.ContentType ;
import org.apache.hc.core5.http.HttpEntity ;
import org.apache.hc.core5.http.io.HttpClientResponseHandler ;
import org.apache.hc.core5.http.io.entity.ByteArrayEntity ;
import org.dom4j.DocumentException ;
import org.dom4j.io.OutputFormat ;
import org.dom4j.io.SAXReader ;
import org.dom4j.io.XMLWriter ;
import org.springframework.beans.factory.annotation.Autowired ;
import org.springframework.http.MediaType ;
import org.springframework.stereotype.Component ;

import com.inzent.yellowpage.marshaller.PublishingFormat ;
import com.inzent.yellowpage.marshaller.YellowPageMarshaller ;
import com.inzent.yellowpage.model.PublishLog ;
import com.inzent.yellowpage.model.PublishTarget ;
import com.inzent.yellowpage.model.Server ;
import com.inzent.yellowpage.model.ServerProperty ;
import com.inzent.yellowpage.model.SystemMeta ;
import com.inzent.yellowpage.model.SystemProperty ;

@Component
public class PublisherBean implements Publisher
{
  protected final Log logger = LogFactory.getLog(getClass()) ;

  @Autowired
  protected PublishingFormatMarshallerBean publishFormatMarshallerBean ;

  @Override
  public String getContextType()
  {
    return MediaType.APPLICATION_XML_VALUE ;
  }

  @Override
  public String getContextString(PublishLog publishLog) throws IOException, DocumentException
  {
    try (StringWriter sw = new StringWriter())
    {
      XMLWriter writer = new XMLWriter(sw, OutputFormat.createPrettyPrint()) ;
      writer.write(SAXReader.createDefault().read(new ByteArrayInputStream(publishLog.getPublishContext()))) ;

      return sw.toString() ;
    }
  }

  @Override
  public byte[] makeContext(PublishLog publishLog) throws Exception
  {
    try (ByteArrayOutputStream baos = new ByteArrayOutputStream(2048))
    {
      YellowPageMarshaller.marshal(baos, publishFormatMarshallerBean.marshal(publishLog)) ;

      return baos.toByteArray() ;
    }
  }

  @Override
  public void publishContext(PublishLog publishLog, PublishTarget publishTarget, Server server)
  {
    // 연계서버의 Property에 등록된 publish.url 의 속성으로 배포요청
    String publishUrl = null ;
    for (ServerProperty serverProperty : server.getServerProperties())
      if (serverProperty.getPk().getKeyId().equals("publish.url"))
      {
        publishUrl = serverProperty.getProperty() ;
        break ;
      }

    if ('Y' == server.getUseYn() && null != publishUrl)
      try
      {
        publishTarget.setUpdateTimestamp(new Timestamp(System.currentTimeMillis())) ;

        HttpUriRequestBase httpRequest = 'Y' == publishLog.getResourceUseYn() ? new HttpPost(publishUrl) : new HttpDelete(publishUrl) ;
        httpRequest.setEntity(new ByteArrayEntity(publishLog.getPublishContext(), ContentType.APPLICATION_XML)) ;

        PublishingFormat publishingFormat = executeHttp(httpRequest, new AbstractHttpClientResponseHandler<PublishingFormat>()
        {
          @Override
          public PublishingFormat handleEntity(HttpEntity entity) throws IOException
          {
            try
            {
              return YellowPageMarshaller.unmarshal(entity.getContent()) ;
            }
            catch (IOException | RuntimeException | Error e)
            {
              throw e ;
            }
            catch (Exception e)
            {
              throw new RuntimeException(e) ;
            }
          }
        }) ;

        // TODO 정상 응답 코드 적용
        if (Objects.equals(publishingFormat.getResultCode(), "0"))
          publishTarget.setPublishDone(PublishTarget.PUBLISH_DONE_SUCCESS) ;
        else
        {
          publishTarget.setPublishDone(PublishTarget.PUBLISH_DONE_ERROR) ;
          publishTarget.setPublishCode(publishingFormat.getResultCode()) ;
          publishTarget.setPublishResult(publishingFormat.getResultMessage()) ;
        }
      }
      catch (Throwable t)
      {
        publishTarget.setPublishDone(PublishTarget.PUBLISH_DONE_ERROR) ;
        publishTarget.setPublishCode(t.getClass().getName()) ;
        publishTarget.setPublishResult(StringUtils.defaultString(t.getMessage(), "null")) ;

        if (logger.isErrorEnabled())
          logger.error(t.getMessage(), t) ;
      }
    else
      publishTarget.setPublishDone(PublishTarget.PUBLISH_DONE_SKIP) ;
  }

  @Override
  public void publishContext(PublishLog publishLog, PublishTarget publishTarget, SystemMeta system)
  {
    // 시스템의 Property에 등록된 publish.url 의 속성으로 배포요청
    String publishUrl = null ;
    for (SystemProperty systemProperty : system.getSystemProperties())
      if (systemProperty.getPk().getKeyId().equals("publish.url"))
      {
        publishUrl = systemProperty.getProperty() ;
        break ;
      }

    if ('Y' == system.getUseYn() && null != publishUrl)
      try
      {
        publishTarget.setUpdateTimestamp(new Timestamp(System.currentTimeMillis())) ;

        HttpUriRequestBase httpRequest = 'Y' == publishLog.getResourceUseYn() ? new HttpPost(publishUrl) : new HttpDelete(publishUrl) ;
        httpRequest.setEntity(new ByteArrayEntity(publishLog.getPublishContext(), ContentType.APPLICATION_XML)) ;

        PublishingFormat publishingFormat = executeHttp(httpRequest, new AbstractHttpClientResponseHandler<PublishingFormat>()
        {
          @Override
          public PublishingFormat handleEntity(HttpEntity entity) throws IOException
          {
            try
            {
              return YellowPageMarshaller.unmarshal(entity.getContent()) ;
            }
            catch (IOException | RuntimeException | Error e)
            {
              throw e ;
            }
            catch (Exception e)
            {
              throw new RuntimeException(e) ;
            }
          }
        }) ;

        // TODO 정상 응답 코드 적용
        if (Objects.equals(publishingFormat.getResultCode(), "0"))
          publishTarget.setPublishDone(PublishTarget.PUBLISH_DONE_SUCCESS) ;
        else
        {
          publishTarget.setPublishDone(PublishTarget.PUBLISH_DONE_ERROR) ;
          publishTarget.setPublishCode(publishingFormat.getResultCode()) ;
          publishTarget.setPublishResult(publishingFormat.getResultMessage()) ;
        }
      }
      catch (Throwable t)
      {
        publishTarget.setPublishDone(PublishTarget.PUBLISH_DONE_ERROR) ;
        publishTarget.setPublishCode(t.getClass().getName()) ;
        publishTarget.setPublishResult(StringUtils.defaultString(t.getMessage(), "null")) ;

        if (logger.isErrorEnabled())
          logger.error(t.getMessage(), t) ;
      }
    else
      publishTarget.setPublishDone(PublishTarget.PUBLISH_DONE_SKIP) ;
  }

  protected PublishingFormat executeHttp(HttpUriRequestBase request, HttpClientResponseHandler<PublishingFormat> responseHandler) throws IOException
  {
    // TODO responseTimeout
    //RequestConfig.Builder builder = RequestConfig.custom() ;
    //builder.setResponseTimeout(Timeout.ofMilliseconds(timeout)) ;
    //request.setConfig(builder.build()) ;

    // TODO connectTimeout
    BasicHttpClientConnectionManager basicHttpClientConnectionManager = new BasicHttpClientConnectionManager() ;
    //basicHttpClientConnectionManager.setConnectionConfig(ConnectionConfig.custom().setConnectTimeout(Timeout.ofMilliseconds(connectTimeout)).build()) ;

    return HttpClientBuilder.create().setConnectionManager(basicHttpClientConnectionManager).build().execute(request, responseHandler) ;
  }
}
