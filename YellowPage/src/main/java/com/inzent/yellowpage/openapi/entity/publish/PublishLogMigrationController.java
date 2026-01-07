/*******************************************************************************
 * This program and the accompanying materials are made
 * available under the terms of the Inzent MCA License v1.0
 * which accompanies this distribution.
 * 
 * Contributors:
 *     Inzent Corporation - initial API and implementation
 *******************************************************************************/
package com.inzent.yellowpage.openapi.entity.publish ;

import java.io.File ;
import java.io.FileInputStream ;
import java.io.FileOutputStream ;
import java.io.IOException ;
import java.io.InputStream ;
import java.io.OutputStream ;
import java.util.HashMap ;
import java.util.Map ;
import java.util.Set ;
import java.util.TreeSet ;

import jakarta.servlet.http.HttpServletRequest ;
import jakarta.servlet.http.HttpServletResponse ;

import org.apache.commons.io.IOUtils ;
import org.springframework.beans.factory.annotation.Autowired ;
import org.springframework.security.core.authority.SimpleGrantedAuthority ;
import org.springframework.security.core.context.SecurityContextHolder ;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping ;
import org.springframework.web.bind.annotation.RequestMapping ;

import com.inzent.imanager.controller.ApiController ;
import com.inzent.imanager.controller.NoPrivilegeException ;
import com.inzent.imanager.marshaller.JsonMarshaller ;
import com.inzent.imanager.openapi.auth.AuthenticationProvider ;
import com.inzent.yellowpage.Privileges ;
import com.inzent.yellowpage.migration.MigrationBean ;
import com.inzent.yellowpage.migration.MigrationList ;
import com.inzent.yellowpage.model.PublishLogPK ;

/**
 * <code>MigrationController</code>
 *
 * @since 2021. 12. 20.
 * @version 5.0
 * @author jaesuk
 */
@Controller
@RequestMapping(PublishLogMigrationController.URI)
public class PublishLogMigrationController extends ApiController
{
  public static final String URI = PublishLogController.URI + "/migration" ;

  public static final String PROT_MIGRATION_SOURCE = "migration.source" ;
  public static final String PROT_MIGRATION_EXPORT = "migration.export" ;
  public static final String PROT_MIGRATION_IMPORT = "migration.import" ;

  @Autowired
  protected MigrationBean migrationBean ;

  protected Set<String> migrationSource = new TreeSet<String>() ;

  public PublishLogMigrationController()
  {
    String migrationTarget = System.getProperty(PROT_MIGRATION_SOURCE) ;
    if (null != migrationTarget)
      for (String src : migrationTarget.split(","))
        migrationSource.add(src.trim()) ;
  }

  @PostMapping("/send")
  public void send(HttpServletRequest request, HttpServletResponse response) throws IOException
  {
    Map<String, Object> model = new HashMap<>() ;

    File file = null ;

    try
    {
      if (!SecurityContextHolder.getContext().getAuthentication().getAuthorities().contains(
          new SimpleGrantedAuthority(AuthenticationProvider.ROLE_PREFIX + Privileges.MIGRATION_EDITOR)))
        throw new NoPrivilegeException() ;

      MigrationList migrationList = JsonMarshaller.unmarshal(IOUtils.toByteArray(request.getInputStream()), MigrationList.class) ;

      File parent = new File(System.getProperty(PROT_MIGRATION_EXPORT, ".")) ;
      parent.mkdirs() ;

      PublishLogPK publishLogPK = migrationList.getPublishLogPKs().get(0) ;
      String name = publishLogPK.getPublishDateTime().replace(':', '-').replace(' ', '_') + "-" + publishLogPK.getPublishId() + ".dat" ;
      file = new File(parent, name) ;

      try (OutputStream os = new FileOutputStream(file)) 
      {
        migrationBean.exportPublishLog(os, migrationList.getPublishLogPKs()) ;
      }

      file = null ;
    }
    catch (Throwable th)
    {
      model.put(MODEL_ERROR, unwrapThrowable(th)) ;
    }

    try
    {
      apiRenderer.renderResponse(request, response, model) ;
    }
    finally
    {
      if (null != file)
        file.delete() ;
    }
  }

  @PostMapping("/recv")
  public void recv(HttpServletRequest request) throws Exception
  {
    if (!migrationSource.contains(request.getRemoteAddr()))
      throw new NoPrivilegeException(request.getRemoteAddr()) ;

    try (InputStream is = new FileInputStream(System.getProperty(PROT_MIGRATION_IMPORT, ".") + File.separator + request.getParameter("import")))
    {
      migrationBean.savePublishLogs(migrationBean.importPublishLog(is)) ;
    }
    catch (Throwable th)
    {
      unwrapThrowable(th) ;

      throw th ;
    }
  }
}
